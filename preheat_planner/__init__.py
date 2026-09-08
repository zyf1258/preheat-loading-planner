"""预热炉排炉核心包的稳定导出入口。

上层服务优先从这里导入公共能力。具体文件仍然可以直接导入，但通过统一入口
可以减少调用方依赖内部目录结构，后续继续拆分时更容易保持兼容。
"""

from .engine import LoadingEngine, allocate_loadings, calculate_preheat_time, iter_placed, normalize_items
from .models import FurnaceSpec, PreheatTimeResult
from .rules import FURNACES
from .service import LoadingPlannerService, normalize_browser_items

__all__ = [
    "FURNACES",
    "FurnaceSpec",
    "LoadingEngine",
    "LoadingPlannerService",
    "PreheatTimeResult",
    "allocate_loadings",
    "calculate_preheat_time",
    "iter_placed",
    "normalize_items",
    "normalize_browser_items",
]
