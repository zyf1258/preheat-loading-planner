"""预热炉排炉核心算法。

本模块只处理三类事情：输入工件校验、按现场规则排炉、预热时间计算。
Excel 文件、网页接口、状态保存和设备控制都属于上层模块，不应混入这里。
这样做的好处是：算法可以被网页、命令行或测试用例复用，而且不需要启动服务器。
"""

from __future__ import annotations

from dataclasses import asdict
import math
from typing import Iterable, Mapping

from .models import FurnaceSpec, PreheatTimeResult
from .rules import (
    BOUND_CODE,
    FURNACES,
    FURNACE_BY_ID,
    LONG_THRESHOLD,
    NIGHT_CODES,
    PREHEAT_DELTA_C,
    PREHEAT_EFFICIENCY,
    PREHEAT_FORMULA_VERSION,
    PREHEAT_LEGACY_FURNACES,
    PREHEAT_OFFSET_THEN_SCALE_FURNACES,
    PREHEAT_SPECIFIC_HEAT_J_PER_KG_C,
    PREHEAT_STANDARD_FURNACES,
    RING_HEIGHT,
    RING_LIMIT,
    normalize_furnace_id,
)


def calculate_preheat_time(furnace_id: object, total_weight_kg: object) -> PreheatTimeResult:
    """根据装炉总质量和炉子功率计算预计预热时间。

    计算分为两步：先根据能量守恒得到理论基础小时数，再套用每类炉子的
    已确认经验修正。1406、1407 仍由现有上层逻辑负责，核心模块只返回警告，
    不会用邻近炉子的参数替代缺失规则。
    """
    normalized_id = normalize_furnace_id(furnace_id)
    spec = FURNACE_BY_ID.get(normalized_id)
    if spec is None:
        return PreheatTimeResult(
            False, normalized_id, None, None, None, None, None, None, "无可用公式",
            ["炉号未配置，不能自动计算预热时间。"],
        )
    try:
        mass = float(total_weight_kg)
    except (TypeError, ValueError):
        mass = 0.0
    if not math.isfinite(mass) or mass <= 0:
        return PreheatTimeResult(
            False, normalized_id, None, spec.power_kw, None, None, None, None, "无可用公式",
            ["装炉总质量缺失或无效，不能自动计算预热时间，请人工设定。"],
        )
    if normalized_id in PREHEAT_LEGACY_FURNACES:
        return PreheatTimeResult(
            False, normalized_id, mass, spec.power_kw, None, None, None, None, "保持现有逻辑",
            [f"{normalized_id} 炉保持现有预热时间逻辑，本公式不覆盖。"],
        )
    if spec.power_kw is None:
        return PreheatTimeResult(
            False, normalized_id, mass, None, None, None, None, None, "功率未配置",
            [f"{normalized_id} 炉子功率未配置，不能自动计算预热时间，请人工设定。"],
        )

    base_hours = (mass * PREHEAT_DELTA_C * PREHEAT_SPECIFIC_HEAT_J_PER_KG_C) / (
        spec.power_kw * 1000.0 * PREHEAT_EFFICIENCY
    ) / 3600.0
    if normalized_id in PREHEAT_OFFSET_THEN_SCALE_FURNACES:
        predicted_hours = (base_hours + 10.0) * 1.3
        adjustment_rule = "(基础预热时间 + 10) × 1.3"
    elif normalized_id in PREHEAT_STANDARD_FURNACES:
        predicted_hours = base_hours * 1.3 + 10.0
        adjustment_rule = "基础预热时间 × 1.3 + 10"
    else:
        return PreheatTimeResult(
            False, normalized_id, mass, spec.power_kw, round(base_hours, 3), None, None, None, "无可用公式",
            [f"{normalized_id} 炉暂无已确认的自动预热时间规则，请人工设定。"],
        )
    return PreheatTimeResult(
        True,
        normalized_id,
        mass,
        spec.power_kw,
        round(base_hours, 3),
        round(predicted_hours, 3),
        adjustment_rule,
        PREHEAT_FORMULA_VERSION,
        "功率-能量公式",
        [],
    )


def _positive_int(value: object, field: str) -> int:
    try:
        number = int(float(str(value).strip()))
    except (AttributeError, TypeError, ValueError) as exc:
        raise ValueError(f"{field} 必须是正整数") from exc
    if number <= 0:
        raise ValueError(f"{field} 必须是正整数")
    return number


def normalize_items(items: Iterable[Mapping[str, object]]) -> list[dict]:
    """校验并复制输入工件，避免排炉过程修改流转程序的原始数据。

    每个输入项需要包含：code、alloy、height_mm、diameter_mm、is_ring、quantity。
    rod_length_mm 可选；未提供时兼容使用 height_mm 作为棒料长度约束值。
    forging_no 可选。quantity 大于 1 时，结果会展开为多个独立炉位工件。
    """
    normalized = []
    for index, raw in enumerate(items):
        if not isinstance(raw, Mapping):
            raise ValueError(f"第 {index + 1} 个工件不是对象")
        code = str(raw.get("code", "")).strip()
        if not code:
            raise ValueError(f"第 {index + 1} 个工件缺少 code")
        alloy = str(raw.get("alloy", "")).strip()
        height_mm = _positive_int(raw.get("height_mm"), "height_mm")
        diameter_mm = _positive_int(raw.get("diameter_mm"), "diameter_mm")
        quantity = _positive_int(raw.get("quantity", 1), "quantity")
        item = {
            "code": code,
            "alloy": alloy,
            "height_mm": height_mm,
            "diameter_mm": diameter_mm,
            "is_ring": bool(raw.get("is_ring", False)),
            "quantity": quantity,
        }
        raw_rod_length = raw.get("rod_length_mm", raw.get("length_mm", raw.get("rod_length")))
        item["rod_length_mm"] = (
            height_mm
            if raw_rod_length in (None, "")
            else _positive_int(raw_rod_length, "rod_length_mm")
        )
        if raw.get("forging_no") not in (None, ""):
            item["forging_no"] = str(raw["forging_no"]).strip()
        normalized.append(item)
    return normalized


class LoadingEngine:
    """按现有预热炉规则进行逐列排炉。

    一个 ``LoadingEngine`` 对象代表“一次完整的自动排炉”。对象内部保存本次
    排炉已经产生的临时状态，包括每台炉的每一列、棒料总长和未分配工件。
    因此同一个对象不要重复调用 ``allocate`` 来排两批互不相关的计划；如果要
    开始新计划，应重新创建一个 ``LoadingEngine``。
    """

    def __init__(self, allowed_furnace_ids: Iterable[object] | None = None) -> None:
        # specs 是本次排炉使用的炉子目录。即使上层只允许部分炉子参与，内部仍
        # 会创建完整结果结构，便于页面显示哪些炉子被排除、哪些炉子为空。
        self.specs = {furnace.furnace_id: furnace for furnace in FURNACES}
        if allowed_furnace_ids is None:
            self.allowed_furnace_ids = tuple(self.specs)
        else:
            if isinstance(allowed_furnace_ids, (str, bytes)):
                raise ValueError("可用炉号必须是炉号列表")
            try:
                values = list(allowed_furnace_ids)
            except TypeError as exc:
                raise ValueError("可用炉号必须是炉号列表") from exc
            normalized = []
            for value in values:
                furnace_id = str(value or "").strip().upper().removeprefix("YJ-")
                if furnace_id not in self.specs:
                    raise ValueError(f"炉号 {value} 不在装炉模块炉号范围内")
                if furnace_id not in normalized:
                    normalized.append(furnace_id)
            self.allowed_furnace_ids = tuple(normalized)
        # 列表用于保持页面和算法的稳定顺序；集合只用于快速判断某台炉是否可用。
        self._allowed_furnace_id_set = set(self.allowed_furnace_ids)
        self.columns = {}
        # 1408 是 950 / 5052 专炉。第一次放入物料时锁定合金族，后续不同族的
        # 物料即使还有空间也不能混入，避免破坏现场的专炉规则。
        self._1408_alloy_family: str | None = None
        for furnace in FURNACES:
            # 1407 对环件使用 ring_slots；其他炉子使用普通 slots。列对象中的
            # items 按炉底到炉顶保存，height_mm 是该列当前累计高度。
            slots = furnace.ring_slots if furnace.furnace_id == "1407" else furnace.slots
            self.columns[furnace.furnace_id] = [
                {"height_mm": 0, "items": []} for _ in range(slots)
            ]
        self.rod_lengths = {furnace.furnace_id: 0 for furnace in FURNACES}
        self.unassigned: list[dict] = []

    def _eligible_candidates(self, candidates: Iterable[str]) -> list[str]:
        """从业务候选顺序中筛出本次实际允许使用的炉子。

        不能直接把 ``allowed_furnace_ids`` 作为候选列表，因为每个业务阶段有
        自己的优先顺序，例如夜班料只能先考虑 1403、1410。
        """

        return [furnace_id for furnace_id in candidates if furnace_id in self._allowed_furnace_id_set]

    @staticmethod
    def _alloy_family(item: Mapping[str, object]) -> str:
        """5052 物料代码或合金描述出现时都视为 5052 族。"""
        return "5052" if "5052" in str(item.get("code", "")) or "5052" in str(item.get("alloy", "")) else "6061"

    @staticmethod
    def _usable_column_count(furnace_id: str, item: Mapping[str, object]) -> int | None:
        if furnace_id != "1407":
            return None
        if not item["is_ring"]:
            return 6
        return 9 if item["code"] == BOUND_CODE else 12

    def _usable_columns(self, furnace_id: str, item: Mapping[str, object]) -> list[dict]:
        columns = self.columns[furnace_id]
        count = self._usable_column_count(furnace_id, item)
        return columns if count is None else columns[:count]

    @staticmethod
    def _best_column_index(columns: list[dict], height: int, max_height: int) -> int | None:
        """选择最适合放入工件的列，并返回列下标。

        选择顺序有三个层次：

        1. 先排除放入后超过炉高上限的列；
        2. 优先选择已经放置件数少的列，让各列层数尽量均匀；
        3. 件数相同再选择当前累计高度较低的列，最后用列号保证结果稳定。
        """

        compatible = [
            index
            for index, column in enumerate(columns)
            if column["height_mm"] + height <= max_height
        ]
        if not compatible:
            return None
        return min(
            compatible,
            key=lambda index: (
                len(columns[index]["items"]),
                columns[index]["height_mm"],
                index,
            ),
        )

    @staticmethod
    def _sort_column(column: dict) -> None:
        # items 按炉底到炉顶保存；高件优先放在底部。
        column["items"].sort(key=lambda placed: -placed["height_mm"])
        height = 0
        for placed in column["items"]:
            height += placed["height_mm"]
            placed["cumulative_mm"] = height
        column["height_mm"] = height

    def _place(
        self,
        furnace_id: str,
        item: Mapping[str, object],
        height: int,
        kind: str,
        limit: int | None = None,
    ) -> bool:
        """尝试把一个工件放进指定炉子的某一列。

        这里集中执行所有“单件放置”时必须满足的硬约束：炉子是否允许使用、
        1408 合金锁、棒料总长、单列高度。只有全部通过后才真正修改内部状态，
        因此返回 ``False`` 时不会留下半条记录。
        """

        if furnace_id not in self._allowed_furnace_id_set:
            return False
        if furnace_id == "1408":
            family = self._alloy_family(item)
            if self._1408_alloy_family and self._1408_alloy_family != family:
                return False
        columns = self._usable_columns(furnace_id, item)
        if not columns:
            return False
        rod_length = item["rod_length_mm"]
        total_limit = self.specs[furnace_id].rod_total_length_mm
        if total_limit is not None and self.rod_lengths[furnace_id] + rod_length > total_limit:
            return False
        max_height = limit if limit is not None else self.specs[furnace_id].height_mm
        index = self._best_column_index(columns, height, max_height)
        if index is None:
            return False
        placed = {
            "code": item["code"],
            "alloy": item["alloy"],
            "height_mm": height,
            "rod_length_mm": rod_length,
            "forging_no": item.get("forging_no"),
            "cumulative_mm": 0,
            "type": kind,
            "is_ring": item["is_ring"],
        }
        columns[index]["items"].append(placed)
        self._sort_column(columns[index])
        self.rod_lengths[furnace_id] += rod_length
        if furnace_id == "1408":
            self._1408_alloy_family = self._alloy_family(item)
        return True

    def _place_preferred(
        self,
        candidates: list[str],
        item: Mapping[str, object],
        height: int,
        kind: str,
        ring: bool = False,
    ) -> bool:
        """按给定炉序尝试放置，适合表达“主炉优先”的业务规则。"""

        for furnace_id in self._eligible_candidates(candidates):
            limit = RING_LIMIT if ring and self.specs[furnace_id].is_ring else None
            if self._place(furnace_id, item, height, kind, limit):
                return True
        return False

    def _place_best(
        self,
        candidates: list[str],
        item: Mapping[str, object],
        height: int,
        kind: str,
        ring: bool = False,
    ) -> bool:
        """在候选炉中选择放入后累计高度最小的一台炉。

        这里的“最优”不是全局数学最优，而是现场排炉规则要求的局部贪心：
        每次模拟把当前工件放到各候选炉的最佳列，比较放入后的列高，选择最小者；
        高度相同则保留候选列表顺序。这样计算简单、结果稳定，也容易解释。
        """

        choices = []
        for candidate_index, furnace_id in enumerate(self._eligible_candidates(candidates)):
            if furnace_id == "1408":
                family = self._alloy_family(item)
                if self._1408_alloy_family and self._1408_alloy_family != family:
                    continue
            columns = self._usable_columns(furnace_id, item)
            if not columns:
                continue
            total_limit = self.specs[furnace_id].rod_total_length_mm
            if total_limit is not None and self.rod_lengths[furnace_id] + item["rod_length_mm"] > total_limit:
                continue
            limit = RING_LIMIT if ring and self.specs[furnace_id].is_ring else self.specs[furnace_id].height_mm
            index = self._best_column_index(columns, height, limit)
            if index is not None:
                new_height = columns[index]["height_mm"] + height
                choices.append((new_height, candidate_index, furnace_id))
        if not choices:
            return False
        _, _, furnace_id = min(choices)
        limit = RING_LIMIT if ring and self.specs[furnace_id].is_ring else None
        return self._place(furnace_id, item, height, kind, limit)

    def _unassigned_reason(self, candidates: Iterable[str], fallback: str) -> str:
        return fallback if self._eligible_candidates(candidates) else "时间轴无可排炉位"

    def _mark_unassigned(self, item: Mapping[str, object], reason: str) -> None:
        self.unassigned.append(
            {
                "code": item["code"],
                "forging_no": item.get("forging_no"),
                "reason": reason,
            }
        )

    def allocate(self, items: Iterable[Mapping[str, object]]) -> dict:
        """执行排炉并返回可序列化的结果字典。

        排炉顺序很重要：先处理限制最强、可选择炉子最少的工件，再处理可使用
        炉子较多的工件，能减少后续出现“有空间但没有合适炉子”的情况。

        当前六个阶段分别是：

        1. 环件进入 1406 / 1407 主炉；
        2. 环件溢出部分尝试进入 1404；
        3. 950 和 5052 料进入 1408，并建立合金锁；
        4. 普通料按当前剩余列高分配；
        5. 夜班料进入 1403 / 1410，必要时补入同族 1408；
        6. 尚未安排的环件使用全炉群剩余空间。
        """
        source_items = normalize_items(items)
        rings = [item for item in source_items if item["is_ring"]]
        normal = [item for item in source_items if not item["is_ring"]]
        night = [item for item in normal if str(item["code"])[-4:] in NIGHT_CODES]
        normal = [item for item in normal if str(item["code"])[-4:] not in NIGHT_CODES]
        ring_overflow = []

        # 1. 环件主炉按炉序尽量装满：1406，再 1407。
        bound = [item for item in rings if item["code"] == BOUND_CODE]
        rest_rings = [item for item in rings if item["code"] != BOUND_CODE]
        for item in bound + sorted(rest_rings, key=lambda value: -value["quantity"]):
            preferred = ["1406"] if item["code"] == BOUND_CODE else ["1406", "1407"]
            for _ in range(item["quantity"]):
                if not self._place_preferred(preferred, item, RING_HEIGHT, "环件", ring=True):
                    ring_overflow.append(item)

        # 2. 1404 共享环件和夜班料的剩余容量。
        ring_units = list(ring_overflow)
        night_units = [item for item in night for _ in range(item["quantity"])]
        remaining_rings, remaining_night = [], list(night_units)
        for item in ring_units:
            if not self._place("1404", item, RING_HEIGHT, "环件", RING_LIMIT):
                remaining_rings.append(item)

        # 3. 1408 接收 950 / 5052，并锁定合金族，禁止 5052 与 6061 混装。
        dedicated = [
            item for item in normal
            if item["diameter_mm"] == 950 or self._alloy_family(item) == "5052"
        ]
        normal = [item for item in normal if item not in dedicated]
        for item in sorted(dedicated, key=lambda value: -value["height_mm"]):
            for _ in range(item["quantity"]):
                if not self._place("1408", item, item["height_mm"], "950 / 5052"):
                    reason = self._unassigned_reason(["1408"], "1408 合金锁定或炉容不足")
                    self._mark_unassigned(item, reason)

        # 4. 普通料选择当前余量最合适的炉列。
        for item in sorted(normal, key=lambda value: -value["height_mm"]):
            candidates = ["1409", "1404", "1405", "1411", "1408"]
            kind = "长料" if item["height_mm"] > LONG_THRESHOLD else "短料"
            for _ in range(item["quantity"]):
                if not self._place_best(candidates, item, item["height_mm"], kind):
                    self._mark_unassigned(item, self._unassigned_reason(candidates, "炉高或炉容不足"))

        # 5. 夜班料只能进入 1403 / 1410；同合金族可补入已锁定的 1408。
        for item in remaining_night:
            candidates = ["1403", "1410"]
            if self._1408_alloy_family == self._alloy_family(item):
                candidates.append("1408")
            if not self._place_best(candidates, item, item["height_mm"], "夜班料"):
                self._mark_unassigned(item, self._unassigned_reason(candidates, "夜班料炉容不足"))

        # 6. 其余环件使用全炉群剩余空间。
        for item in remaining_rings:
            candidates = ["1409", "1410", "1411", "1403", "1405", "1408", "1406", "1404"]
            if item["code"] != BOUND_CODE:
                candidates.insert(6, "1407")
            if not self._place_best(candidates, item, RING_HEIGHT, "环件填充", ring=True):
                self._mark_unassigned(item, self._unassigned_reason(candidates, "全炉群容量不足"))

        return self.result(source_items)

    def result(self, source_items: list[dict]) -> dict:
        """把内部临时状态整理成接口和页面可以直接消费的字典。

        ``total_requested`` 按输入数量统计，``total_placed`` 按实际生成的炉位
        统计，二者之差由 ``unassigned`` 解释。每台炉同时返回棒料总长使用量和
        剩余量，方便上层页面展示约束结果。
        """

        furnaces = {}
        placed = 0
        for furnace_id, columns in self.columns.items():
            spec = self.specs[furnace_id]
            count = sum(len(column["items"]) for column in columns)
            placed += count
            furnaces[furnace_id] = {
                "spec": asdict(spec),
                "total": count,
                "columns": columns,
                "max_height_mm": max((column["height_mm"] for column in columns), default=0),
                "used_rod_length_mm": self.rod_lengths[furnace_id],
                "remaining_rod_length_mm": (
                    None
                    if spec.rod_total_length_mm is None
                    else spec.rod_total_length_mm - self.rod_lengths[furnace_id]
                ),
            }
        return {
            "total_requested": sum(item["quantity"] for item in source_items),
            "total_placed": placed,
            "total_unassigned": len(self.unassigned),
            "allowed_furnace_ids": list(self.allowed_furnace_ids),
            "excluded_furnace_ids": [
                furnace_id for furnace_id in self.specs if furnace_id not in self._allowed_furnace_id_set
            ],
            "furnaces": furnaces,
            "unassigned": self.unassigned,
        }


def allocate_loadings(
    items: Iterable[Mapping[str, object]],
    available_furnace_ids: Iterable[object] | None = None,
) -> dict:
    """模块推荐入口：传入工件列表，返回排炉结果。"""
    return LoadingEngine(available_furnace_ids).allocate(items)


def iter_placed(result: Mapping[str, object]):
    """按炉号、列号、炉内顺序遍历已放置工件。"""
    furnaces = result.get("furnaces", {})
    for furnace_id, furnace in furnaces.items():
        for column_index, column in enumerate(furnace.get("columns", []), start=1):
            for item_index, item in enumerate(column.get("items", []), start=1):
                yield {
                    "furnace_id": furnace_id,
                    "column": column_index,
                    "position": item_index,
                    **item,
                }


__all__ = [
    "BOUND_CODE",
    "FURNACES",
    "FurnaceSpec",
    "LONG_THRESHOLD",
    "LoadingEngine",
    "NIGHT_CODES",
    "PREHEAT_FORMULA_VERSION",
    "PreheatTimeResult",
    "RING_HEIGHT",
    "RING_LIMIT",
    "allocate_loadings",
    "calculate_preheat_time",
    "iter_placed",
    "normalize_items",
]
