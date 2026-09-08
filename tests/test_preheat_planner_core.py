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


if __name__ == "__main__":
    unittest.main()
