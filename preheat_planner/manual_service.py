"""装炉结果的手动调整服务。

自动排炉由 ``engine.py`` 负责，而用户拖拽换炉属于另一类操作：它接收一个
已经存在的结果字典，移动其中的一件工件，然后重新计算列高和虚拟炉摘要。
本模块只处理这种结果维护，不重新执行自动排炉。
"""

from __future__ import annotations

import copy
from collections.abc import Mapping

from .rules import RING_LIMIT


VIRTUAL_FURNACE = "VIRTUAL"
VIRTUAL_FURNACE_Z = "VIRTUAL_Z"
UNCONFIGURED_ENTITY_FURNACES = ("1412", "1413")


def _number(value: object) -> int:
    """把页面传来的数字或数字字符串转换为整数。"""

    try:
        return int(float(str(value or "").strip() or 0))
    except (TypeError, ValueError):
        return 0


def refresh_column(column: dict) -> None:
    """按炉底到炉顶重新排序一列，并更新累计高度。"""

    column["items"].sort(key=lambda item: -item["height_mm"])
    height = 0
    for item in column["items"]:
        height += item["height_mm"]
        item["cumulative_mm"] = height
    column["height_mm"] = height


def add_unconfigured_entity_furnaces(result: dict) -> dict:
    """向结果补充 1412、1413 的展示占位，但不赋予自动排炉能力。"""

    furnaces = result.setdefault("furnaces", {})
    for furnace_id in UNCONFIGURED_ENTITY_FURNACES:
        furnaces.setdefault(
            furnace_id,
            {
                "spec": {
                    "furnace_id": furnace_id,
                    "slots": 0,
                    "height_mm": 0,
                    "speed": "",
                    "coefficient": 0,
                    "base_heat_h": 0,
                    "role": "实体炉（炉体规格待配置）",
                    "is_large": False,
                    "is_ring": False,
                    "ring_slots": 0,
                    "is_virtual": False,
                    "loading_configured": False,
                },
                "total": 0,
                "columns": [],
                "max_height_mm": 0,
            },
        )
    return result


def refresh_virtual_summary(result: dict) -> None:
    """根据虚拟炉内容刷新未分配、已装入数量和各炉摘要。"""

    virtual_items = result["furnaces"][VIRTUAL_FURNACE]["columns"][0]["items"]
    result["unassigned"] = [
        {
            "code": item["code"],
            "forging_no": item.get("forging_no"),
            "reason": item.get("reason", "未分配"),
        }
        for item in virtual_items
    ]
    result["total_unassigned"] = len(virtual_items)
    result["total_placed"] = sum(
        furnace["total"]
        for furnace_id, furnace in result["furnaces"].items()
        if furnace_id not in (VIRTUAL_FURNACE, VIRTUAL_FURNACE_Z)
    )


def add_virtual_furnace(result: dict, source_items: list[dict]) -> dict:
    """创建待分配虚拟炉和手动暂存炉。"""

    # 缺长度或容量不足的记录只保留了 code/锻造号/reason，因此先从原始工件
    # 找回合金、长度和环件标志，保证页面重新拖拽时信息完整。
    by_identity = {(item["code"], item.get("forging_no")): item for item in source_items}
    virtual_items = []
    for entry in result.get("unassigned", []):
        source = by_identity.get((entry["code"], entry.get("forging_no")), {})
        quantity = _number(entry.get("quantity")) or 1
        for _ in range(quantity):
            virtual_items.append(
                {
                    "code": entry["code"],
                    "forging_no": entry.get("forging_no"),
                    "alloy": source.get("alloy", ""),
                    "height_mm": source.get("height_mm", 0),
                    "cumulative_mm": 0,
                    "type": "待分配",
                    "is_ring": source.get("is_ring", False),
                    "reason": entry.get("reason", "未分配"),
                }
            )

    result.setdefault("furnaces", {})[VIRTUAL_FURNACE] = {
        "spec": {
            "furnace_id": VIRTUAL_FURNACE,
            "slots": 0,
            "height_mm": 0,
            "speed": "",
            "coefficient": 0,
            "base_heat_h": 0,
            "role": "未分配暂存（最低优先级）",
            "is_large": False,
            "is_ring": False,
            "ring_slots": 0,
            "is_virtual": True,
        },
        "total": len(virtual_items),
        "columns": [{"height_mm": 0, "items": virtual_items}],
        "max_height_mm": 0,
    }
    result["furnaces"][VIRTUAL_FURNACE_Z] = {
        "spec": {
            "furnace_id": VIRTUAL_FURNACE_Z,
            "slots": 0,
            "height_mm": 0,
            "speed": "",
            "coefficient": 0,
            "base_heat_h": 0,
            "role": "手动暂存，不参与自动排炉",
            "is_large": False,
            "is_ring": False,
            "ring_slots": 0,
            "is_virtual": True,
            "manual_only": True,
        },
        "total": 0,
        "columns": [{"height_mm": 0, "items": []}],
        "max_height_mm": 0,
    }
    refresh_virtual_summary(result)
    return result


def _recalculate_summaries(result: dict) -> None:
    """重新计算所有炉子的总件数和最高列高度。"""

    for furnace in result["furnaces"].values():
        furnace["total"] = sum(len(column.get("items", [])) for column in furnace.get("columns", []))
        furnace["max_height_mm"] = max(
            (column.get("height_mm", 0) for column in furnace.get("columns", [])),
            default=0,
        )


def move_item(
    result: dict,
    source_furnace: object,
    source_column: object,
    item_index: object,
    target_furnace: object,
    target_column: object | None = None,
) -> dict:
    """移动一件工件，并执行目标炉的高度约束。"""

    source_furnace = str(source_furnace)
    target_furnace = str(target_furnace)
    source_column = _number(source_column)
    item_index = _number(item_index)
    furnaces = result["furnaces"]
    if source_furnace not in furnaces or target_furnace not in furnaces:
        raise ValueError("炉号不存在")
    source_columns = furnaces[source_furnace]["columns"]
    if not 0 <= source_column < len(source_columns):
        raise ValueError("来源列不存在")
    source = source_columns[source_column]
    if not 0 <= item_index < len(source["items"]):
        raise ValueError("来源棒料不存在")
    item = source["items"][item_index]
    target = furnaces[target_furnace]

    if target_furnace in (VIRTUAL_FURNACE, VIRTUAL_FURNACE_Z):
        source["items"].pop(item_index)
        refresh_column(source)
        target["columns"][0]["items"].append(item)
        _recalculate_summaries(result)
        refresh_virtual_summary(result)
        return result

    if item.get("height_mm", 0) <= 0:
        raise ValueError("请先补充该物料的棒料长度，再拖入实体炉")
    target_spec = target.get("spec") or {}
    limit = (
        RING_LIMIT
        if item.get("is_ring") and target_spec.get("is_ring")
        else target_spec.get("height_mm", 0)
    )
    target_columns = target.get("columns") or []
    if target_column is not None:
        target_index = _number(target_column)
        if not 0 <= target_index < len(target_columns):
            raise ValueError(f"{target_furnace} 第 {target_index + 1} 列不存在")
        if target_columns[target_index]["height_mm"] + item["height_mm"] > limit:
            raise ValueError(f"{target_furnace} 第 {target_index + 1} 列高度不足")
    else:
        candidates = [
            (column["height_mm"], index)
            for index, column in enumerate(target_columns)
            if column["height_mm"] + item["height_mm"] <= limit
        ]
        if not candidates:
            raise ValueError(f"{target_furnace} 没有满足高度约束的空位")
        _, target_index = min(candidates)

    source["items"].pop(item_index)
    refresh_column(source)
    target_columns[target_index]["items"].append(item)
    refresh_column(target_columns[target_index])
    _recalculate_summaries(result)
    refresh_virtual_summary(result)
    return result


def batch_move(result: dict, moves: object, target_furnace: object) -> dict:
    """按稳定索引顺序批量换炉，并返回移动后的结果。"""

    if not isinstance(moves, list) or not moves or not target_furnace:
        raise ValueError("请选择至少一根棒料和目标炉")
    result = copy.deepcopy(result)
    normalized = [
        {
            "source_furnace": move.get("source_furnace", move.get("furnace")),
            "source_column": move.get("source_column", move.get("column")),
            "item_index": move.get("item_index", move.get("item")),
        }
        for move in moves
        if isinstance(move, Mapping)
    ]
    if len(normalized) != len(moves) or any(not move["source_furnace"] for move in normalized):
        raise ValueError("批量换炉缺少来源炉信息")
    ordered = sorted(
        normalized,
        key=lambda move: (
            str(move["source_furnace"]),
            _number(move["source_column"]),
            -_number(move["item_index"]),
        ),
    )
    for move in ordered:
        move_item(
            result,
            move["source_furnace"],
            move["source_column"],
            move["item_index"],
            target_furnace,
        )
    return result


__all__ = [
    "UNCONFIGURED_ENTITY_FURNACES",
    "VIRTUAL_FURNACE",
    "VIRTUAL_FURNACE_Z",
    "add_unconfigured_entity_furnaces",
    "add_virtual_furnace",
    "batch_move",
    "move_item",
    "refresh_column",
    "refresh_virtual_summary",
]
