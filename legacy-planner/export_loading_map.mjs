import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const [inputPath, outputPath, templatePath] = process.argv.slice(2);
const result = JSON.parse(await fs.readFile(inputPath, 'utf8'));
const template = await FileBlob.load(templatePath);
const workbook = await SpreadsheetFile.importXlsx(template);
const sheet = workbook.worksheets.getItem('预热炉装炉图');

// V5 arranges each furnace as groups of three physical columns. Each entry is
// one furnace column with its three Excel cells listed from top layer to bottom.
const threeColumnGroups = (columns, groups) => groups.flatMap(rows =>
  columns.map(column => rows.map(row => `${column}${row}`)));
const layouts = {
  '1408': { title: 'A4', cells: [['A6', 'A7', 'A8'], ['B6', 'B7', 'B8'], ['C6', 'C7', 'C8'], ['A10', 'A11', 'A12']] },
  '1405': { title: 'D4', cells: threeColumnGroups(['D', 'E', 'F'], [[6, 7, 8], [10, 11, 12]]) },
  '1409': { title: 'G4', cells: threeColumnGroups(['G', 'H', 'I'], [[6, 7, 8], [10, 11, 12]]) },
  '1404': { title: 'J4', cells: threeColumnGroups(['J', 'K', 'L'], [[6, 7, 8], [10, 11, 12]]) },
  '1403': { title: 'A13', cells: threeColumnGroups(['A', 'B', 'C'], [[15, 16, 17], [19, 20, 21]]) },
  '1406': { title: 'D13', cells: threeColumnGroups(['D', 'E', 'F'], [[15, 16, 17], [19, 20, 21]]) },
  '1407': { title: 'H13', cells: threeColumnGroups(['H', 'I', 'J', 'K'], [[15, 16, 17], [19, 20, 21], [23, 24, 25]]) },
  '1410': { title: 'A22', cells: threeColumnGroups(['A', 'B', 'C'], [[25, 26, 27], [29, 30, 31]]) },
  '1411': { title: 'D22', cells: threeColumnGroups(['D', 'E', 'F'], [[25, 26, 27], [29, 30, 31]]) },
};
const gridOverflow = [];

sheet.getRange('A3').values = [[`导出时间：${new Date().toLocaleString('zh-CN', { hour12: false })}`]];
sheet.getRange('G3').values = [[`计划数量：${result.total_requested}　已装入：${result.total_placed}`]];

for (const [furnaceId, layout] of Object.entries(layouts)) {
  const furnace = result.furnaces[furnaceId];
  if (!furnace) continue;
  const title = `${furnaceId}　${furnace.spec.role}　炉高上限 ${furnace.spec.height_mm}mm`;
  sheet.getRange(layout.title).values = [[title]];

  for (let columnIndex = 0; columnIndex < layout.cells.length; columnIndex += 1) {
    const items = furnace.columns[columnIndex]?.items || [];
    const columnCells = layout.cells[columnIndex];
    for (const cell of columnCells) sheet.getRange(cell).clear({ applyTo: 'contents' });
    for (let itemIndex = 0; itemIndex < Math.min(items.length, columnCells.length); itemIndex += 1) {
      // Source stacks are bottom-to-top. Excel rows are written from the bottom upward.
      const target = sheet.getRange(columnCells[columnCells.length - 1 - itemIndex]);
      const item = items[itemIndex];
      target.values = [[`${item.forging_no || item.code}\n${item.height_mm}mm`]];
      target.format = {
        fill: item.is_ring ? '#FDE9B4' : '#DDEEF4',
        wrapText: true, horizontalAlignment: 'center', verticalAlignment: 'center',
        borders: { preset: 'all', style: 'thin', color: '#777777' },
      };
    }
    for (let itemIndex = columnCells.length; itemIndex < items.length; itemIndex += 1) {
      const item = items[itemIndex];
      gridOverflow.push({
        ...item,
        furnace_id: furnaceId,
        column_index: columnIndex,
        layer_index: itemIndex + 1,
      });
    }
  }
}

const virtualItems = result.furnaces?.VIRTUAL?.columns?.[0]?.items
  ?? (result.unassigned || []).map(item => ({
    code: item.code,
    forging_no: item.forging_no,
    height_mm: item.height_mm,
  }));
const overflowStartRow = 38;
sheet.getRange(`A${overflowStartRow}:C400`).clear({ applyTo: 'contents' });
let unassignedStartRow = overflowStartRow;
if (gridOverflow.length) {
  sheet.getRange(`A${overflowStartRow}`).values = [[`第四层备注（${gridOverflow.length} 根）`]];
  sheet.getRange(`A${overflowStartRow}`).format = {
    fill: '#FFF2CC', font: { bold: true }, horizontalAlignment: 'left',
  };
  const overflowHeaderRow = overflowStartRow + 1;
  sheet.getRange(`A${overflowHeaderRow}:C${overflowHeaderRow}`).values = [[
    '炉号', '列', '备注',
  ]];
  sheet.getRange(`A${overflowHeaderRow}:C${overflowHeaderRow}`).format = {
    fill: '#FFF2CC', font: { bold: true }, horizontalAlignment: 'center',
    borders: { preset: 'all', style: 'thin', color: '#777777' },
  };
  const overflowRows = gridOverflow.map(item => {
    const label = item.forging_no
      ? `${item.forging_no}（${item.code}）`
      : (item.code || '未命名工件');
    const lengthText = item.height_mm ? `${item.height_mm}mm` : '待补充长度';
    return [item.furnace_id, `第${item.column_index + 1}列`,
      `第${item.layer_index}层：${label} / ${lengthText}`];
  });
  const overflowDataStart = overflowHeaderRow + 1;
  const overflowDataEnd = overflowDataStart + overflowRows.length - 1;
  sheet.getRange(`A${overflowDataStart}:C${overflowDataEnd}`).values = overflowRows;
  sheet.getRange(`A${overflowDataStart}:C${overflowDataEnd}`).format = {
    fill: '#FFFBEA', horizontalAlignment: 'center', verticalAlignment: 'center',
    borders: { preset: 'all', style: 'thin', color: '#777777' },
  };
  unassignedStartRow = overflowDataEnd + 2;
}

const unassigned = [...virtualItems];
sheet.getRange(`A${unassignedStartRow}`).values = [[`未排炉工件（${unassigned.length} 根）`]];
sheet.getRange(`A${unassignedStartRow}`).format = {
  fill: '#FCE4D6', font: { bold: true }, horizontalAlignment: 'left',
};
sheet.getRange(`A${unassignedStartRow + 1}:C${unassignedStartRow + 1}`).values = [[
  '物料编码', '锻造号', '棒料长度',
]];
sheet.getRange(`A${unassignedStartRow + 1}:C${unassignedStartRow + 1}`).format = {
  fill: '#FCE4D6', font: { bold: true }, horizontalAlignment: 'center',
  borders: { preset: 'all', style: 'thin', color: '#777777' },
};
const unassignedRows = unassigned.length
  ? unassigned.map(item => [item.code || '', item.forging_no || item.code || '',
    item.height_mm ? `${item.height_mm}mm` : '待补充'])
  : [['无', '', '']];
const unassignedDataStart = unassignedStartRow + 2;
const unassignedDataEnd = unassignedDataStart + unassignedRows.length - 1;
sheet.getRange(`A${unassignedDataStart}:C${unassignedDataEnd}`).values = unassignedRows;
sheet.getRange(`A${unassignedDataStart}:C${unassignedDataEnd}`).format = {
  fill: '#FFF7ED', horizontalAlignment: 'center', verticalAlignment: 'center',
  borders: { preset: 'all', style: 'thin', color: '#777777' },
};

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
