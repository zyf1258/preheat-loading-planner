"""预热炉核心模块的单元测试。

这些测试不启动 Web 服务，也不依赖 Excel 模板，专门验证第一阶段拆分后的核心
模块仍然保持输入、输出和旧核心一致。这样后续继续拆 Excel 读写时，算法有独立
的安全网。
"""

from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path

from preheat_planner.engine import allocate_loadings, calculate_preheat_time, iter_placed, normalize_items
from preheat_planner.service import DEFAULT_SERVICE, normalize_browser_items
from preheat_planner.rules import FURNACE_BY_ID, normalize_furnace_id


ROOT = Path(__file__).parents[1]


def _load_reference_module():
    """加载重构前保留的核心副本，用于做结果对比。"""

    module_path = ROOT / "preheat-furnace-module" / "preheat_furnace_module.py"
    spec = importlib.util.spec_from_file_location("reference_preheat_furnace_module", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载参考核心：{module_path}")
    module = importlib.util.module_from_spec(spec)
    # dataclass 在 Python 3.12 中会通过 sys.modules 查找类所属模块。
    # 手动加载文件时要先注册模块，行为才和正常 import 一致。
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _ring_item(code: str, quantity: int, rod_length_mm: int = 430) -> dict:
    return {
        "code": code,
        "alloy": "6061",
        "height_mm": 430,
        "diameter_mm": 900,
        "rod_length_mm": rod_length_mm,
        "is_ring": True,
        "quantity": quantity,
    }


class PreheatPlannerCoreTest(unittest.TestCase):
    def test_furnace_id_normalization_keeps_one_canonical_form(self):
        self.assertEqual(normalize_furnace_id(" yj-1405 "), "1405")
        self.assertIs(FURNACE_BY_ID["1405"], FURNACE_BY_ID.get(normalize_furnace_id("YJ-1405")))

    def test_normalize_items_copies_input_and_uses_height_as_length_fallback(self):
        source = {
            "code": "F3.6061.0201",
            "alloy": "6061",
            "height_mm": "703",
            "diameter_mm": "630",
            "quantity": "2",
            "is_ring": False,
        }

        normalized = normalize_items([source])

        self.assertEqual(normalized[0]["height_mm"], 703)
        self.assertEqual(normalized[0]["rod_length_mm"], 703)
        self.assertEqual(normalized[0]["quantity"], 2)
        self.assertNotIn("rod_length_mm", source)

    def test_refactored_engine_matches_reference_core(self):
        items = [
            {
                "code": "F2.6061.0322",
                "alloy": "6061",
                "height_mm": 430,
                "diameter_mm": 800,
                "rod_length_mm": 430,
                "is_ring": True,
                "quantity": 3,
                "forging_no": "D28945-B-1",
            },
            {
                "code": "F3.5052.0100",
                "alloy": "5052",
                "height_mm": 500,
                "diameter_mm": 700,
                "rod_length_mm": 1200,
                "is_ring": False,
                "quantity": 2,
                "forging_no": "D30001-A-1",
            },
            {
                "code": "F3.6061.0121",
                "alloy": "6061",
                "height_mm": 600,
                "diameter_mm": 630,
                "rod_length_mm": 600,
                "is_ring": False,
                "quantity": 1,
                "forging_no": "D30002-A-1",
            },
        ]
        reference = _load_reference_module()

        self.assertEqual(
            allocate_loadings(items),
            reference.allocate_loadings(items),
        )

    def test_preheat_time_matches_reference_for_each_configured_furnace(self):
        reference = _load_reference_module()

        for furnace_id in FURNACE_BY_ID:
            with self.subTest(furnace_id=furnace_id):
                self.assertEqual(
                    calculate_preheat_time(furnace_id, 1000).as_dict(),
                    reference.calculate_preheat_time(furnace_id, 1000).as_dict(),
                )

    def test_iter_placed_contains_furnace_and_column_location(self):
        result = allocate_loadings(
            [
                {
                    "code": "F3.6061.0201",
                    "alloy": "6061",
                    "height_mm": 700,
                    "diameter_mm": 630,
                    "rod_length_mm": 900,
                    "is_ring": False,
                    "quantity": 1,
                }
            ],
            available_furnace_ids=["1409"],
        )

        placed = list(iter_placed(result))

        self.assertEqual(len(placed), 1)
        self.assertEqual(placed[0]["furnace_id"], "1409")
        self.assertEqual(placed[0]["column"], 1)
        self.assertEqual(placed[0]["position"], 1)
        self.assertEqual(placed[0]["rod_length_mm"], 900)

    def test_service_normalizes_historical_browser_field_names(self):
        browser_item = {
            "material_code": "F3.6061.0201",
            "workpiece_type": "锻件",
            "height_mm": 700,
            "diameter_mm": 630,
            "length_mm": 900,
            "quantity": 1,
        }

        normalized = normalize_browser_items([browser_item])
        result = DEFAULT_SERVICE.allocate([browser_item], ["1409"])

        self.assertEqual(normalized[0]["code"], "F3.6061.0201")
        self.assertEqual(normalized[0]["rod_length_mm"], 900)
        self.assertEqual(result["total_placed"], 1)
        self.assertEqual(result["furnaces"]["1409"]["used_rod_length_mm"], 900)


class RingAllocationRuleTest(unittest.TestCase):
    def _furnace_total(self, result: dict, furnace_id: str) -> int:
        return sum(len(column["items"]) for column in result["furnaces"][furnace_id]["columns"])

    def test_large_ring_only_enters_1406_with_quantity_cap(self):
        # 大环是 F2.6061.0001。棒料总长已按 15 × 430mm 放开，因此按标准
        # 430mm 棒料排炉时，数量上限 15 个能够真正达到。
        result = allocate_loadings([_ring_item("F2.6061.0001", 17)])

        self.assertEqual(self._furnace_total(result, "1406"), 15)
        self.assertEqual(result["total_unassigned"], 2)
        self.assertEqual(
            [entry["reason"] for entry in result["unassigned"]],
            ["1406 大环容量不足", "1406 大环容量不足"],
        )
        other_totals = {
            furnace_id: self._furnace_total(result, furnace_id)
            for furnace_id in result["furnaces"]
            if furnace_id != "1406"
        }
        self.assertEqual(sum(other_totals.values()), 0)

    def test_1406_only_holds_large_rings(self):
        # 1406 只装大环 F2.6061.0001；小环（含 F1.6061.0001）只进 1407。
        result = allocate_loadings([
            _ring_item("F2.6061.0001", 3, rod_length_mm=360),
            _ring_item("F2.6061.0322", 5),
            _ring_item("F1.6061.0001", 2),
        ])

        self.assertEqual(self._furnace_total(result, "1406"), 3)
        self.assertEqual(self._furnace_total(result, "1407"), 7)
        self.assertEqual(result["total_unassigned"], 0)

    def test_1407_holds_thirty_six_conventional_rings(self):
        # 1407 是 12 列 × 3 层，放开棒料总长后正好容纳 36 根普通环件。
        result = allocate_loadings(
            [_ring_item("F2.6061.0322", 36)],
            available_furnace_ids=["1407"],
        )

        self.assertEqual(self._furnace_total(result, "1407"), 36)
        self.assertEqual(result["total_unassigned"], 0)

    def test_remaining_rings_no_longer_fill_other_furnaces(self):
        # 小环只使用 1407 / 1404（最多 36 + 18 = 54 根）；超出的 1 根应未分配，
        # 既不进入只装大环的 1406，也不按旧的“环件填充”规则进入 1405。
        result = allocate_loadings(
            [_ring_item("F2.6061.0322", 55)],
            available_furnace_ids=["1406", "1407", "1404", "1405"],
        )

        self.assertEqual(result["total_placed"], 54)
        self.assertEqual(result["total_unassigned"], 1)
        self.assertEqual(self._furnace_total(result, "1406"), 0)
        self.assertEqual(self._furnace_total(result, "1405"), 0)
        self.assertEqual(
            {entry["reason"] for entry in result["unassigned"]},
            {"环件炉容量不足"},
        )


if __name__ == "__main__":
    unittest.main()
