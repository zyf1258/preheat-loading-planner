"""排炉应用服务层。

核心算法只认识规范化后的工件字典，而浏览器提交的数据可能使用
``material_code``、``workpiece_type`` 等页面字段。这个服务层负责把页面输入
转换成核心输入，并把核心结果原样交给上层，不处理 HTTP 和 Excel。
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping

from .engine import allocate_loadings as _allocate_loadings
from .engine import calculate_preheat_time as _calculate_preheat_time
from .rules import FURNACES


def _as_bool(value: object) -> bool:
    """把浏览器常见的真假值转换成布尔值。"""

    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in {
        "1", "true", "yes", "y", "on", "是", "环件", "碾环", "冲孔",
    }


def supported_furnace_ids() -> tuple[str, ...]:
    """返回新版核心支持的炉号目录。"""

    return tuple(spec.furnace_id for spec in FURNACES)


def normalize_browser_items(items: object) -> list[dict]:
    """把浏览器行数据转换成排炉核心需要的统一字段。

    页面历史版本同时使用过 ``code`` 和 ``material_code``，也使用过布尔值、
    中文工艺名称表示环件。兼容逻辑集中在这里，算法层不再需要了解页面字段。
    """

    if not isinstance(items, list) or not items:
        raise ValueError("至少需要一件下一炉工件")

    normalized = []
    for index, raw in enumerate(items, start=1):
        if not isinstance(raw, Mapping):
            raise ValueError(f"第 {index} 行工件不是对象")
        workpiece_type = str(raw.get("workpiece_type") or "").strip()
        alloy = str(raw.get("alloy") or workpiece_type or "锻件").strip()
        normalized.append(
            {
                "code": raw.get("code", raw.get("material_code", "")),
                "alloy": alloy,
                "height_mm": raw.get("height_mm", ""),
                "diameter_mm": raw.get("diameter_mm", ""),
                "rod_length_mm": raw.get(
                    "rod_length_mm",
                    raw.get("length_mm", raw.get("rod_length", "")),
                ),
                "is_ring": _as_bool(raw.get("is_ring")) or workpiece_type in {"碾环", "冲孔", "锻造碾环"},
                "quantity": raw.get("quantity", 1),
                "forging_no": raw.get("forging_no", ""),
            }
        )
    return normalized


class LoadingPlannerService:
    """面向应用层的排炉服务。

    服务对象本身不保存计划状态；每次调用都创建独立的核心排炉过程，避免一次
    请求的数据污染下一次请求。可用炉号由调用方传入，便于未来接入炉况系统。
    """

    def allocate(self, items: object, available_furnace_ids: Iterable[object] | None = None) -> dict:
        """校验页面输入并执行自动排炉。"""

        normalized = normalize_browser_items(items)
        return _allocate_loadings(normalized, available_furnace_ids)

    def calculate_preheat_time(self, furnace_id: object, total_weight_kg: object) -> dict:
        """计算预热时间并转换为接口字典。"""

        return _calculate_preheat_time(furnace_id, total_weight_kg).as_dict()

    def empty_result(self, available_furnace_ids: Iterable[object] | None = None) -> dict:
        """创建空排炉结果，供导入装炉图等恢复炉结构的流程使用。"""

        return _allocate_loadings([], available_furnace_ids)


DEFAULT_SERVICE = LoadingPlannerService()


__all__ = [
    "DEFAULT_SERVICE",
    "LoadingPlannerService",
    "normalize_browser_items",
    "supported_furnace_ids",
]
