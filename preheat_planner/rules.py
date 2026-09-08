"""预热炉排炉使用的现场规则、炉子目录和规则辅助函数。

这里集中保存不会随着某一次排炉过程变化的配置。排炉过程中的临时状态，例如
每列已经放了多少工件，仍然由 ``engine.LoadingEngine`` 自己管理。
"""

from __future__ import annotations

from .models import FurnaceSpec


# 环件在排炉算法中按 430 mm 的固定层高处理；环件填充炉最多使用三层。
RING_HEIGHT = 430
RING_LIMIT = RING_HEIGHT * 3

# 超过该高度的普通工件只用于标记“长料”，实际可排炉范围仍由炉体高度校验。
LONG_THRESHOLD = 1000

# 这个物料编码在 1407 炉有单独的有效列数规则。
BOUND_CODE = "F2.6061.0001"

# 物料编码末四位属于夜班料标识时，只能优先进入 1403 / 1410。
NIGHT_CODES = ("0121", "0199", "0436")

# 预热时间模型的常量。公式含义是：
# 基础小时数 = 质量 × 升温温差 × 比热 / (功率 × 效率) / 3600。
PREHEAT_FORMULA_VERSION = "power-energy-v1"
PREHEAT_DELTA_C = 470.0
PREHEAT_SPECIFIC_HEAT_J_PER_KG_C = 904.0
PREHEAT_EFFICIENCY = 0.35

# 不同炉子使用不同的经验修正规则。
PREHEAT_STANDARD_FURNACES = frozenset({"1405", "1408", "1409", "1410"})
PREHEAT_OFFSET_THEN_SCALE_FURNACES = frozenset({"1403", "1404", "1411"})
PREHEAT_LEGACY_FURNACES = frozenset({"1406", "1407"})


# 炉子目录是有顺序的。排炉算法会按这个顺序生成结果和初始化炉列，不能随意改成无序集合。
FURNACES = (
    FurnaceSpec("1403", 6, 1600, "慢", 1.5, 14.0, "夜班炉", rod_total_length_mm=7800, power_kw=800),
    FurnaceSpec("1404", 6, 1600, "慢", 1.5, 12.0, "早班炉", is_ring=True, rod_total_length_mm=7800, power_kw=800),
    FurnaceSpec("1405", 6, 1900, "慢", 1.5, 9.4, "中班炉", rod_total_length_mm=7300, power_kw=1000),
    FurnaceSpec("1406", 6, 1600, "中等", 1.35, 7.0, "环件主炉", is_ring=True, rod_total_length_mm=5600, power_kw=900),
    FurnaceSpec(
        "1407",
        6,
        1750,
        "快",
        1.25,
        9.0,
        "环件主炉（普通环件12根/层；F2.6061.0001为9根/层）",
        is_ring=True,
        ring_slots=12,
        rod_total_length_mm=9200,
        power_kw=1000,
    ),
    FurnaceSpec("1408", 4, 2200, "快", 1.25, 8.0, "950 / 5052 专炉", is_large=True, rod_total_length_mm=8700, power_kw=1200),
    FurnaceSpec("1409", 6, 2200, "快", 1.25, 14.0, "早班炉", is_large=True, rod_total_length_mm=13900, power_kw=1400),
    FurnaceSpec("1410", 6, 2000, "中等", 1.35, 11.0, "夜班炉", is_large=True, rod_total_length_mm=11700, power_kw=1200),
    FurnaceSpec("1411", 6, 2100, "中等", 1.35, 11.0, "普通炉", is_large=True, power_kw=1380),
)

# 字典只读语义由模块本身保证：调用方拿到的是副本时才允许修改。
FURNACE_BY_ID = {spec.furnace_id: spec for spec in FURNACES}


def normalize_furnace_id(value: object) -> str:
    """统一炉号写法，兼容 ``1405`` 和 ``YJ-1405`` 两种输入。"""

    return str(value or "").strip().upper().removeprefix("YJ-")


__all__ = [
    "BOUND_CODE",
    "FURNACES",
    "FURNACE_BY_ID",
    "LONG_THRESHOLD",
    "NIGHT_CODES",
    "PREHEAT_DELTA_C",
    "PREHEAT_EFFICIENCY",
    "PREHEAT_FORMULA_VERSION",
    "PREHEAT_LEGACY_FURNACES",
    "PREHEAT_OFFSET_THEN_SCALE_FURNACES",
    "PREHEAT_SPECIFIC_HEAT_J_PER_KG_C",
    "PREHEAT_STANDARD_FURNACES",
    "RING_HEIGHT",
    "RING_LIMIT",
    "normalize_furnace_id",
]
