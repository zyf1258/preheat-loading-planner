"""预热炉排炉核心使用的数据模型。

这个文件只描述数据长什么样，不负责读取 Excel、处理 HTTP 请求或执行排炉。
把数据模型独立出来后，学习和维护排炉规则时可以先看清楚输入、输出的结构，
不会被具体的算法步骤干扰。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass(frozen=True)
class FurnaceSpec:
    """一台预热炉的静态规格。

    字段含义：

    * ``furnace_id``：炉号，例如 ``1405``；统一使用不带 ``YJ-`` 前缀的文本。
    * ``slots``：常规情况下的列数。
    * ``height_mm``：单列允许的最大高度。
    * ``speed``、``coefficient``、``base_heat_h``、``role``：现场规则和展示信息。
    * ``is_large``：是否属于大炉，供页面或上层业务展示使用。
    * ``is_ring``：是否是环件专用或主要用于环件的炉子。
    * ``ring_slots``：环件排炉时的有效列数；为 0 时使用 ``slots``。
    * ``rod_total_length_mm``：整台炉允许的棒料总长上限。
    * ``power_kw``：预热时间公式使用的额定功率；缺失时不能自动计算时间。
    """

    furnace_id: str
    slots: int
    height_mm: int
    speed: str
    coefficient: float
    base_heat_h: float
    role: str
    is_large: bool = False
    is_ring: bool = False
    ring_slots: int = 0
    rod_total_length_mm: int | None = None
    power_kw: float | None = None


@dataclass(frozen=True)
class PreheatTimeResult:
    """预热时间计算结果。

    ``available`` 用来区分“已经有可信自动结果”和“需要人工设定”两种状态。
    因此调用方不应该只判断 ``predicted_hours`` 是否为空，还应同时检查该字段。
    ``warnings`` 保留给页面展示明确原因，例如炉号未配置、重量无效或炉子沿用旧逻辑。
    """

    available: bool
    furnace_id: str
    total_weight_kg: float | None
    power_kw: float | None
    base_hours: float | None
    predicted_hours: float | None
    adjustment_rule: str | None
    formula_version: str | None
    model_label: str
    warnings: list[str]

    def as_dict(self) -> dict:
        """转换成适合 JSON 接口返回的普通字典。"""

        return asdict(self)


__all__ = ["FurnaceSpec", "PreheatTimeResult"]
