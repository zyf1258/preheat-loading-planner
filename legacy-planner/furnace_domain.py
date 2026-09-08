"""纯排炉领域模块：不包含时间线、炉况或现场控制。"""
from dataclasses import asdict, dataclass

RING_HEIGHT = 430
RING_LIMIT = RING_HEIGHT * 3
LONG_THRESHOLD = 1000
BOUND_CODE = 'F2.6061.0001'
NIGHT_CODES = ('0121', '0199', '0436')


@dataclass(frozen=True)
class FurnaceSpec:
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


FURNACES = (
    FurnaceSpec('1403', 6, 1600, '慢', 1.5, 14.0, '夜班炉'),
    FurnaceSpec('1404', 6, 1600, '慢', 1.5, 12.0, '早班炉', is_ring=True),
    FurnaceSpec('1405', 6, 1900, '慢', 1.5, 9.4, '中班炉'),
    FurnaceSpec('1406', 6, 1600, '中等', 1.35, 7.0, '环件主炉', is_ring=True),
    FurnaceSpec('1407', 6, 1750, '快', 1.25, 9.0,
                '环件主炉（普通环件12根/层；F2.6061.0001为9根/层）', is_ring=True, ring_slots=12),
    FurnaceSpec('1408', 4, 2200, '快', 1.25, 8.0, '中班炉 / 950 / 5052 专炉', is_large=True),
    FurnaceSpec('1409', 6, 2200, '快', 1.25, 14.0, '早班炉', is_large=True),
    FurnaceSpec('1410', 6, 2000, '中等', 1.35, 11.0, '夜班炉', is_large=True),
    FurnaceSpec('1411', 6, 2100, '中等', 1.35, 11.0, '普通炉', is_large=True),
)


class LoadingEngine:
    """保留旧系统的核心优先级与逐列高度感知放置规则。"""
    def __init__(self):
        self.specs = {f.furnace_id: f for f in FURNACES}
        self.columns = {}
        self._1408_alloy_family = None
        for spec in FURNACES:
            slots = spec.ring_slots if spec.furnace_id == '1407' else spec.slots
            self.columns[spec.furnace_id] = [{'height_mm': 0, 'items': []} for _ in range(slots)]
        self.unassigned = []

    @staticmethod
    def _alloy_family(item):
        """5052 可出现在物料代码或合金描述中，二者都必须参与隔离判定。"""
        return '5052' if '5052' in item['code'] or '5052' in item['alloy'] else '6061'

    @staticmethod
    def _usable_column_count(furnace_id, item):
        if furnace_id != '1407':
            return None
        if not item['is_ring']:
            return 6
        # F2.6061.0001 is the large-ring family and keeps its original 9-up layout.
        return 9 if item['code'] == BOUND_CODE else 12

    def _usable_columns(self, furnace_id, item):
        columns = self.columns[furnace_id]
        count = self._usable_column_count(furnace_id, item)
        return columns if count is None else columns[:count]

    @staticmethod
    def _best_column_index(columns, height, max_height):
        """Complete the lowest item-count layer before starting the next one."""
        compatible = [
            index for index, column in enumerate(columns)
            if column['height_mm'] + height <= max_height
        ]
        if not compatible:
            return None
        return min(compatible, key=lambda index: (
            len(columns[index]['items']), columns[index]['height_mm'], index
        ))

    @staticmethod
    def _sort_column(column):
        # The list represents the stack from furnace bottom to top.
        column['items'].sort(key=lambda placed: -placed['height_mm'])
        height = 0
        for placed in column['items']:
            height += placed['height_mm']
            placed['cumulative_mm'] = height
        column['height_mm'] = height

    def _place(self, furnace_id, item, height, kind, limit=None):
        if furnace_id == '1408':
            family = self._alloy_family(item)
            if self._1408_alloy_family and self._1408_alloy_family != family:
                return False
        columns = self._usable_columns(furnace_id, item)
        if not columns:
            return False
        max_height = limit if limit is not None else self.specs[furnace_id].height_mm
        index = self._best_column_index(columns, height, max_height)
        if index is None:
            return False
        columns[index]['items'].append({
            'code': item['code'], 'alloy': item['alloy'], 'height_mm': height,
            'forging_no': item.get('forging_no'),
            'cumulative_mm': 0, 'type': kind,
            'is_ring': item['is_ring'],
        })
        self._sort_column(columns[index])
        if furnace_id == '1408':
            self._1408_alloy_family = self._alloy_family(item)
        return True

    def _place_preferred(self, candidates, item, height, kind, ring=False):
        """按炉序尽量填满，用于环件主炉。"""
        for furnace_id in candidates:
            limit = RING_LIMIT if ring and self.specs[furnace_id].is_ring else None
            if self._place(furnace_id, item, height, kind, limit):
                return True
        return False

    def _place_best(self, candidates, item, height, kind, ring=False):
        choices = []
        for candidate_index, furnace_id in enumerate(candidates):
            if furnace_id == '1408':
                family = self._alloy_family(item)
                if self._1408_alloy_family and self._1408_alloy_family != family:
                    continue
            cols = self._usable_columns(furnace_id, item)
            if not cols:
                continue
            limit = RING_LIMIT if ring and self.specs[furnace_id].is_ring else self.specs[furnace_id].height_mm
            index = self._best_column_index(cols, height, limit)
            if index is not None:
                new_height = cols[index]['height_mm'] + height
                # Keep the business candidate order when loads are equal.
                # Otherwise tuple sorting falls back to furnace-id text order.
                choices.append((new_height, candidate_index, furnace_id))
        if not choices:
            return False
        _, _, furnace_id = min(choices)
        return self._place(furnace_id, item, height, kind, RING_LIMIT if ring and self.specs[furnace_id].is_ring else None)

    def allocate(self, items):
        rings = [x for x in items if x['is_ring']]
        normal = [x for x in items if not x['is_ring']]
        night = [x for x in normal if x['code'][-4:] in NIGHT_CODES]
        normal = [x for x in normal if x['code'][-4:] not in NIGHT_CODES]
        ring_overflow = []

        # 1. 环件主炉按炉序尽量装满：1406，再 1407；1404 与夜班料共享。
        bound = [x for x in rings if x['code'] == BOUND_CODE]
        rest_rings = [x for x in rings if x['code'] != BOUND_CODE]
        for item in bound + sorted(rest_rings, key=lambda x: -x['quantity']):
            # 1407 is reserved for small rings during automatic allocation.
            preferred = ['1406'] if item['code'] == BOUND_CODE else ['1406', '1407']
            for _ in range(item['quantity']):
                if not self._place_preferred(preferred, item, RING_HEIGHT, '环件', ring=True):
                    ring_overflow.append(item)

        # 2. 1404 是环件与夜班量产料的同级共享炉：两类料交替占用其容量。
        ring_units = list(ring_overflow)
        night_units = [item for item in night for _ in range(item['quantity'])]
        shared = [(item, RING_HEIGHT, '环件') for item in ring_units]
        remaining_rings, remaining_night = [], list(night_units)
        for item, fixed_height, kind in shared:
            height = fixed_height or item['height_mm']
            if not self._place('1404', item, height, kind, RING_LIMIT if fixed_height else None):
                (remaining_rings if fixed_height else remaining_night).append(item)

        # 3. 1408 先接收 5052 / 950，写入合金锁，禁止 5052 与 6061 混装。
        dedicated = [x for x in normal if x['diameter_mm'] == 950 or self._alloy_family(x) == '5052']
        normal = [x for x in normal if x not in dedicated]
        for item in sorted(dedicated, key=lambda x: -x['height_mm']):
            for _ in range(item['quantity']):
                if not self._place('1408', item, item['height_mm'], '950 / 5052'):
                    self.unassigned.append({'code': item['code'], 'forging_no': item.get('forging_no'),
                                            'reason': '1408 合金锁定或炉容不足'})

        # 4. 其它非环件：环件优先排完后，1404/1406/1407 的剩余空间可继续使用。
        for item in sorted(normal, key=lambda x: -x['height_mm']):
            dedicated = item['diameter_mm'] == 950 or self._alloy_family(item) == '5052'
            candidates = ['1408'] if dedicated else (
                ['1409', '1404', '1405', '1411', '1408']
                if item['height_mm'] > LONG_THRESHOLD else
                ['1409', '1404', '1405', '1411', '1408']
            )
            kind = '950 / 5052' if dedicated else ('长料' if item['height_mm'] > LONG_THRESHOLD else '短料')
            for _ in range(item['quantity']):
                if not self._place_best(candidates, item, item['height_mm'], kind):
                    self.unassigned.append({'code': item['code'], 'forging_no': item.get('forging_no'),
                                            'reason': '炉高或炉容不足'})

        # 5. 夜班料只能进入夜班炉；若 1408 已形成合金锁，允许同合金族夜班料补充其余量。
        for item in remaining_night:
            night_candidates = ['1403', '1410']
            if self._1408_alloy_family == self._alloy_family(item):
                night_candidates.append('1408')
            if not self._place_best(night_candidates, item, item['height_mm'], '夜班料'):
                self.unassigned.append({'code': item['code'], 'forging_no': item.get('forging_no'),
                                        'reason': '夜班料炉容不足'})
        for item in remaining_rings:
            all_candidates = ['1409', '1410', '1411', '1403', '1405', '1408', '1406', '1404']
            if item['code'] != BOUND_CODE:
                all_candidates.insert(6, '1407')
            if not self._place_best(all_candidates, item, RING_HEIGHT, '环件填充', ring=True):
                self.unassigned.append({'code': item['code'], 'forging_no': item.get('forging_no'),
                                        'reason': '全炉群容量不足'})
        return self.result(items)

    def result(self, source_items):
        furnaces = {}
        placed = 0
        for fid, columns in self.columns.items():
            spec = self.specs[fid]
            count = sum(len(c['items']) for c in columns)
            placed += count
            furnaces[fid] = {
                'spec': asdict(spec), 'total': count, 'columns': columns,
                'max_height_mm': max(c['height_mm'] for c in columns),
            }
        return {
            'total_requested': sum(x['quantity'] for x in source_items),
            'total_placed': placed,
            'total_unassigned': len(self.unassigned),
            'furnaces': furnaces, 'unassigned': self.unassigned,
        }
