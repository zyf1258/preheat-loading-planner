# 预热炉排炉工作台（独立版）

这是从连续预热炉看板中拆出的独立排炉项目，只包含排炉相关能力：

- 导入当天生产计划和 MES 棒料登记表；
- 自动排炉、炉高和棒料总长约束校验；
- 手动添加、拖放和换炉；
- 棒料长度记忆；
- 预热时间显示；
- 装炉图导出及右侧装炉明细。

本项目不连接连续看板、不读取 SQLite 数据库，也不修改原项目和历史源码。1412、1413 作为实体炉显示，但因炉体规格尚未配置，暂不参与自动排炉；原有虚拟炉保持原逻辑。

## 代码结构

新版核心位于 `preheat_planner` 包：

- `models.py`：炉子规格和预热时间结果模型；
- `rules.py`：炉号目录、容量约束和现场规则常量；
- `engine.py`：唯一自动排炉算法和预热时间计算；
- `service.py`：浏览器字段兼容和应用服务编排；
- `excel_io.py`：生产计划、MES 棒料表读取；
- `import_io.py`：装炉图导入恢复；
- `manual_service.py`：手动换炉、虚拟炉维护和实体炉占位；
- `export_io.py`：装炉图模板导出、明细合并和预估重量。

页面由 `app.py` 提供：`static/loading-plan.html` 是独立版外壳，内嵌
`/legacy-planner.html` 渲染的排炉界面；`legacy_adapter.py` 是该界面与
`preheat_planner` 核心之间的桥接层，同时沿用旧目录的长度记忆、解析器和
导出模板。因此 `legacy-planner/` 是运行时依赖（旧页面、长度记忆和 Excel
模板），不是可删除的历史包袱；`preheat-furnace-module/` 仅作为测试中的
参考实现，用于校验新旧算法结果一致。

## 运行

需要 Python 3.10 或更高版本，并安装依赖：

```powershell
python -m pip install -r requirements.txt
python app.py
```

默认地址：`http://127.0.0.1:8770/`

Windows 可双击 `start-planner.bat` 启动。可通过环境变量修改端口：

```powershell
$env:PREHEAT_PLANNER_PORT = "8770"
python app.py
```

## 数据位置

默认保存在项目目录的 `data` 文件夹中：

- `billet_lengths.json`：棒料长度记忆。装炉安排保存在浏览器会话中，
  通过导出的装炉图 Excel 留档，服务端不再持久化。

部署到权限受限目录时，建议把数据目录指定到可写位置：

```powershell
$env:PREHEAT_PLANNER_DATA_ROOT = "D:\PreheatFurnacePlannerData"
python app.py
```

## 公司内网部署准备

1. 在服务器安装 Python 3.10+。
2. 将本目录完整复制到服务器，不要只复制 `app.py`。
3. 执行 `python -m pip install -r requirements.txt`。
4. 设置 `PREHEAT_PLANNER_DATA_ROOT` 到服务器备份目录。
5. 运行 `start-planner.bat`，或使用 Windows 任务计划/服务管理器启动 `python app.py`。
6. 防火墙只放行公司内网需要访问的端口；默认监听本机地址，如需局域网访问，再由反向代理或明确的内网绑定方式提供。
7. 定期备份 `billet_lengths.json` 和 `legacy-planner/assets/` 下的 Excel 模板。

## 接口自检

打开：`http://127.0.0.1:8770/api/v1/health`

返回 `ok: true` 即表示独立排炉服务已启动。

## 测试

```powershell
python -m unittest discover -s tests -v
```
