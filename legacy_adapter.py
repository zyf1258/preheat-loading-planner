"""旧排程页面与独立服务之间的兼容适配层。

这个文件是独立项目和原有“预热炉排炉”页面之间的桥梁：

* 独立项目的 HTTP 服务在 :mod:`app` 中；
* 原排程页面、排炉算法和 Excel 模板仍保存在 ``legacy-planner`` 目录；
* 本文件负责动态加载旧模块、把上传文件转换为旧模块能理解的结构，
  再把旧模块的结果转换回独立页面和导出文件需要的结构。

学习这个文件时，可以把它看成三层：

1. **输入层**：读取 Excel、长度记忆和浏览器提交的换炉参数；
2. **兼容层**：隔离旧页面、长度记忆和历史接口格式，同时接入新版核心服务；
3. **输出层**：整理页面结果，并把结果写回装炉图模板。

本文件中的大多数函数不会重新发明业务规则，而是负责“适配数据格式”。
因此，阅读时要特别留意每个函数的输入字典字段和输出字典字段。
"""
from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
from pathlib import Path

from preheat_planner.excel_io import load_plan_items, read_billet_length_map
from preheat_planner.export_io import (
    export_furnace_details as _new_export_furnace_details,
    export_loading_map_with_openpyxl as _new_export_loading_map_with_openpyxl,
)
from preheat_planner.import_io import restore_loading_map as _new_restore_loading_map
from preheat_planner.manual_service import (
    add_unconfigured_entity_furnaces as _new_add_unconfigured_entity_furnaces,
    add_virtual_furnace as _new_add_virtual_furnace,
    batch_move as _new_batch_move,
    move_item as _new_move_item,
)
from preheat_planner.service import DEFAULT_SERVICE


ROOT = Path(__file__).resolve().parent
# 原项目拆分后的根目录。默认就在当前独立项目目录下，也可以用环境变量
# LEGACY_PREHEAT_ROOT 指向另一份旧排程模块。
DEFAULT_LEGACY_ROOT = ROOT / "legacy-planner"
LEGACY_ROOT = Path(os.environ.get(
    "LEGACY_PREHEAT_ROOT",
    DEFAULT_LEGACY_ROOT,
)).resolve()
# 旧页面的 HTML 文件。页面中的 /api/ 路径会在 render_legacy_page() 中被
# 改写为 /api/legacy/，从而经过当前服务的兼容接口。
LEGACY_INDEX = LEGACY_ROOT / "static" / "index.html"


def _read_billet_length_map(path: Path, legacy_app) -> dict:
    """兼容旧函数名，转发到独立 Excel 读取层。"""
    return read_billet_length_map(path, legacy_app.number)


def _load_legacy_items(path: Path, lengths: dict, legacy_app) -> tuple[list[dict], list[dict]]:
    """兼容旧函数名，转发到独立 Excel 计划读取层。"""
    return load_plan_items(
        path,
        lengths,
        ring_height=legacy_app.RING_HEIGHT,
        number_parser=legacy_app.number,
        forging_parser=legacy_app.parse_forging_numbers,
    )


def _legacy_app():
    """动态加载并缓存旧排程模块。

    ``legacy-planner`` 不是标准 Python 包名目录，且旧模块依赖同目录下的
    ``furnace_domain.py``。因此这里使用 ``importlib`` 指定文件路径加载，
    并在执行模块代码前临时把旧目录放到 ``sys.path``。

    这是一个“按需加载 + 进程内缓存”的函数：第一次调用时加载模块，之后直接
    返回 ``sys.modules`` 中的同一个对象，避免重复加载、重复读取旧模块状态。
    """
    module_name = "preheat_furnace_legacy_app"
    loaded = sys.modules.get(module_name)
    if loaded is not None:
        # 动态模块已经加载过，直接复用。
        return loaded
    app_file = LEGACY_ROOT / "app.py"
    if not app_file.is_file():
        raise FileNotFoundError(f"找不到旧排程模块：{app_file}")
    spec = importlib.util.spec_from_file_location(module_name, app_file)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载旧排程模块：{app_file}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    legacy_path = str(LEGACY_ROOT)
    # 旧 app.py 的 import 语句需要在自身目录中寻找 furnace_domain.py。
    sys.path.insert(0, legacy_path)
    try:
        spec.loader.exec_module(module)
    except Exception:
        # 加载失败时清理半成品模块，避免下一次调用拿到不完整对象。
        sys.modules.pop(module_name, None)
        raise
    finally:
        # 无论成功还是失败，都撤销临时 sys.path 修改，避免污染整个进程。
        if sys.path and sys.path[0] == legacy_path:
            sys.path.pop(0)
    return module


def read_lengths() -> dict:
    """对外暴露旧模块的长度记忆读取接口。

    页面和当前服务不需要知道长度文件具体存在哪里，因此通过这个小包装函数
    隔离旧模块实现细节。
    """
    return _legacy_app().read_lengths()


def remember_length(code: object, length_mm: object) -> dict:
    """校验并保存用户手动补充的棒料长度。

    浏览器传来的 JSON 值可能是字符串、整数或空值，所以先统一转成字符串和
    数字，再执行大于 0 的校验。保存仍委托给旧模块，以保持原有长度文件格式。
    """
    app = _legacy_app()
    code_value = str(code or "").strip()
    length_value = app.number(length_mm)
    if not code_value or length_value <= 0:
        raise ValueError("请输入物料编码和大于 0 的长度")
    lengths = app.read_lengths()
    lengths[code_value] = length_value
    app.save_lengths(lengths)
    return {"ok": True, "code": code_value, "length_mm": length_value}


def generate_loading_plan(plan_bytes: bytes, billet_bytes: bytes | None = None) -> dict:
    """生成一次完整的自动排炉结果。

    处理顺序如下：

    1. 确认上传了生产计划；
    2. 读取已有长度记忆；
    3. 把上传内容写入临时目录，必要时导入 MES 长度并更新记忆；
    4. 解析生产计划，分出有效项和缺长度项；
    5. 调用新版 ``preheat_planner`` 核心执行排炉；
    6. 把缺长度项和未分配项放进虚拟炉，返回给页面。

    临时文件只用于让旧的 openpyxl 读取逻辑复用上传表格，函数结束后会自动删除。
    """
    app = _legacy_app()
    if not plan_bytes:
        raise ValueError("请选择当天生产计划")
    lengths = app.read_lengths()
    with tempfile.TemporaryDirectory() as temp:
        # 旧模块的读取函数需要文件路径，因此先把 HTTP 上传的 bytes 落盘。
        plan_path = Path(temp) / "plan.xlsx"
        plan_path.write_bytes(plan_bytes)
        if billet_bytes:
            # MES 棒料表是可选的：首次导入或有新料时上传即可。
            billet_path = Path(temp) / "billet.xlsx"
            billet_path.write_bytes(billet_bytes)
            lengths.update(_read_billet_length_map(billet_path, app))
            app.save_lengths(lengths)
        items, missing = _load_legacy_items(plan_path, lengths, app)
    # 只把有效项交给自动排炉。缺长度项稍后合并回结果，避免数据丢失。
    # 旧模块在这里只负责页面兼容能力和旧长度记忆；真正的自动排炉及虚拟炉
    # 维护已经交给新版核心服务，页面不再依赖 legacy-planner 的旧算法。
    result = DEFAULT_SERVICE.allocate(items)
    result["unassigned"].extend(missing)
    result["total_unassigned"] = len(result["unassigned"])
    # 给尚未配置的实体炉增加展示信息，但不改变自动算法已产生的装炉结果。
    _new_add_unconfigured_entity_furnaces(result)
    # 旧模块的 unassigned 列表还不能直接被页面拖拽，需要转成虚拟炉结构。
    _new_add_virtual_furnace(result, items)
    return {"ok": True, "items": items, "result": result, "remembered_lengths": len(lengths)}


def import_loading_map(map_bytes: bytes) -> dict:
    """兼容旧接口，转发到新版装炉图恢复服务。

    导入解析、左侧炉位恢复、锻造号范围展开和超层工件处理都已经集中到
    ``preheat_planner.import_io``，适配层只保留原有的函数签名，保证旧页面
    的请求地址和返回结构不需要同步改动。
    """

    return _new_restore_loading_map(map_bytes, DEFAULT_SERVICE.empty_result)


def move_item(body: dict) -> dict:
    """兼容旧接口，转发到新版单件换炉服务。"""

    result = _new_move_item(
        body["result"],
        body["source_furnace"],
        body["source_column"],
        body["item_index"],
        body["target_furnace"],
        body.get("target_column"),
    )
    return {"ok": True, "result": result}


def batch_move(body: dict) -> dict:
    """兼容旧接口，转发到新版批量换炉服务。"""

    result = _new_batch_move(
        body["result"],
        body.get("moves", []),
        body.get("target_furnace"),
    )
    return {"ok": True, "result": result, "moved": len(body.get("moves", []))}


def _export_furnace_details(result: dict) -> list[list[object]]:
    """兼容旧函数名，转发到新版 Excel 导出层。"""
    return _new_export_furnace_details(result)


def export_loading_map_with_openpyxl(result: dict, template_path: Path) -> bytes:
    """兼容旧函数名，转发到新版 Excel 导出层。"""
    return _new_export_loading_map_with_openpyxl(result, template_path)


def export_loading_map(result: dict) -> bytes:
    """导出统一入口：检查模板存在后执行 Excel 生成并包装异常。"""
    app = _legacy_app()
    if not app.EXPORT_TEMPLATE.exists():
        raise ValueError("未找到装炉图导出模板")
    try:
        return export_loading_map_with_openpyxl(result, app.EXPORT_TEMPLATE)
    except Exception as exc:
        raise ValueError(f"装炉图导出失败：{exc}") from exc


def render_legacy_page() -> str:
    """读取旧排程页面并注入独立项目所需的兼容脚本。

    页面原本请求 ``/api/...``，而独立服务需要把旧接口放在
    ``/api/legacy/...`` 下，因此先替换请求前缀。随后插入一小段桥接脚本，
    用来显示尚未配置的 1412、1413 实体炉。

    这里返回 HTML 字符串，不直接写文件；HTTP 服务收到请求时会把字符串返回给
    浏览器，这样旧页面源文件本身保持不变。
    """
    if not LEGACY_INDEX.is_file():
        raise FileNotFoundError(f"找不到旧排程页面：{LEGACY_INDEX}")
    html = LEGACY_INDEX.read_text(encoding="utf-8")
    # 同时处理单引号和双引号写法，兼容旧页面中不同版本的 fetch 调用。
    html = html.replace("fetch('/api/", "fetch('/api/legacy/")
    html = html.replace('fetch("/api/', 'fetch("/api/legacy/')
    bridge = r'''<script>
(() => {
  // 1412、1413 尚未配置装炉尺寸，只展示身份，不参与自动排炉。
  const entityFurnaces = ['1412', '1413'];
  function renderEntityFurnaces() {
    // 旧页面每次 render 都会重绘炉区，所以先删除可能过期的同炉节点。
    const grid = document.getElementById('furnaces');
    if (!grid) return;
    entityFurnaces.forEach(id => {
      // 如果旧页面将来自己提供了该炉，就不额外插入桥接节点。
      grid.querySelector('[data-target="' + id + '"]')?.remove();
      if (grid.querySelector('[data-legacy-entity="' + id + '"]')) return;
      const panel = document.createElement('article');
      panel.className = 'panel legacy-entity-furnace';
      panel.dataset.legacyEntity = id;
      panel.innerHTML = '<div class="furnace-head"><h2>YJ-' + id + '</h2><b>实体炉</b></div><div class="meta">实体炉 · 炉体规格待配置 · 暂不参与自动排炉</div><div class="items"><div class="item">暂无装入物料</div></div>';
      const virtual = grid.querySelector('[data-target="VIRTUAL"]');
      // 虚拟炉应保持在实体炉之后，因此优先插入到虚拟炉之前。
      if (virtual) grid.insertBefore(panel, virtual); else grid.appendChild(panel);
    });
  }
  const originalRender = window.render;
  if (typeof originalRender === 'function') {
    // 包装原 render，在旧页面重绘完成后补回实体炉展示节点。
    window.render = function() {
      const result = originalRender.apply(this, arguments);
      setTimeout(renderEntityFurnaces, 0);
      return result;
    };
  }
  // 定时补偿旧页面异步渲染或第三方代码覆盖 render 的情况。
  setInterval(renderEntityFurnaces, 500);
})();
</script>'''
    return html.replace("</body>", bridge + "</body>")
