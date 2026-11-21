import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const [input, directory, mode] = process.argv.slice(2);
if (!input || !directory) throw new Error('Cần đường dẫn JSON và thư mục đầu ra mới');
const modules = process.env.REPORT_NODE_MODULES;
const require = modules
  ? createRequire(pathToFileURL(join(resolve(modules), '..', 'report-runtime.cjs')))
  : createRequire(import.meta.url);
const { Workbook, SpreadsheetFile } = await import(pathToFileURL(require.resolve('@oai/artifact-tool')));
const { Document, Packer, Paragraph, HeadingLevel } = require('docx');
const data = JSON.parse(await readFile(input, 'utf8'));
const prefix = `Tuần ${data.week}`;
const missing = 'Thiếu dữ liệu';

const hqHeaders = ['Đơn vị', 'PTM (đ)', 'TSD (đ)', 'Sản lượng PTM',
  'PRO · Số điểm', 'PRO · Điểm PSDT', 'PRO · DT đánh giá (đ)', 'PRO · Điểm PSSL',
  'PRO · SL PTM', 'PRO · Chi lương', 'KEY1 · Số điểm', 'KEY1 · Điểm PSDT',
  'KEY1 · DT đánh giá (đ)', 'KEY1 · Điểm PSSL', 'KEY1 · SL PTM', 'KEY1 · Chi lương'];
const bscHeaders = ['Đơn vị', 'Nhân viên', 'NV đạt 20%', 'KH I.8.1', 'HT I.8.1',
  'Điểm I.8.1', 'Tổng điểm bán', 'Điểm PSSL', 'KH I.8.2', 'HT I.8.2',
  'Điểm I.8.2', 'KH DT (triệu đ)', 'DT (triệu đ)', 'HT I.8.3', 'Điểm I.8.3', 'BSC chung'];

function sheet(book, name, headers, rows, options = {}) {
  const tab = book.worksheets.add(name);
  tab.showGridLines = false;
  tab.freezePanes.freezeRows(headers.length);
  const width = headers[0].length;
  const matrix = [...headers, ...rows].map(r => Array.from({ length: width }, (_, i) => r[i] ?? null));
  tab.getRangeByIndexes(0, 0, matrix.length, width).values = matrix;
  const head = tab.getRangeByIndexes(0, 0, headers.length, width);
  head.format.fill = '#173554';
  head.format.font.color = '#FFFFFF';
  head.format.font.bold = true;
  head.format.wrapText = true;
  head.format.rowHeight = options.headerHeight ?? 40;
  for (let c = 0; c < width; c++) {
    tab.getRangeByIndexes(0, c, matrix.length, 1).format.columnWidthPx = options.widths?.[c] ?? 112;
  }
  if (rows.length) {
    const body = tab.getRangeByIndexes(headers.length, 0, rows.length, width);
    body.format.numberFormat = '#,##0.00';
    const totals = tab.getRangeByIndexes(matrix.length - 1, 0, 1, width);
    if (rows.at(-1)[0] === 'Toàn tỉnh') {
      totals.format.fill = '#DBEAF6';
      totals.format.font.bold = true;
    }
    for (const c of options.percent ?? []) {
      tab.getRangeByIndexes(headers.length, c, rows.length, 1).format.numberFormat = '0.0%';
    }
    for (const c of options.integer ?? []) {
      tab.getRangeByIndexes(headers.length, c, rows.length, 1).format.numberFormat = '#,##0';
    }
  }
  return tab;
}

function ratio(n, d) {
  return d ? n / d : null;
}

function weeklyHQ(row, index) {
  const [unit, ptm, reuse, prod, pro, proRevPoints, proRev, proActive, proQty, proPay,
    vk, vkRevPoints, vkRev, vkActive, vkQty, vkPay] = row;
  const transactionPTM = data.program[index][3];
  return [unit, proActive + vkActive, transactionPTM + reuse, reuse, transactionPTM, prod, reuse,
    pro, proRevPoints, proRev, ratio(proRevPoints, pro), proActive, proQty, ratio(proActive, pro),
    proPay, missing, vk, vkRevPoints, vkRev, ratio(vkRevPoints, vk),
    vkActive, vkQty, ratio(vkActive, vk), vkPay, missing, ...Array(7).fill('Bổ sung sau')];
}

function weeklyProgram(row, index) {
  const [unit, pro, vk, sum] = row;
  const cells = Array(43).fill(missing);
  cells[0] = unit === 'Toàn tỉnh' ? null : index + 1;
  cells[1] = unit;
  cells[14] = pro / 1_000_000;
  cells[15] = vk / 1_000_000;
  cells[16] = sum / 1_000_000;
  return cells;
}

function printSheet(book, name, title, headers, rows, options) {
  const tab = sheet(book, name, [[title, ...headers.slice(1).map(() => null)], headers], rows, options);
  const width = headers.length;
  const cover = tab.getRangeByIndexes(0, 0, 1, width);
  cover.merge();
  cover.format.rowHeight = 32;
  cover.format.font.size = 16;
  const labels = tab.getRangeByIndexes(1, 0, 1, width);
  labels.format.fill = '#214D68';
  labels.format.rowHeight = 42;
  for (const [i, row] of rows.entries()) {
    const band = tab.getRangeByIndexes(i + 2, 0, 1, width);
    band.format.rowHeight = 21;
    if (row[0] !== 'Toàn tỉnh' && i % 2 === 1) band.format.fill = '#EDF3F7';
    tab.getRangeByIndexes(i + 2, 0, 1, 1).format.font.bold = true;
  }
}

function addHqPrintSheets(book) {
  printSheet(book, 'In doanh thu', `Doanh thu kênh bán hàng · ${prefix}`,
    ['Đơn vị', 'PTM theo giao dịch', 'PTM theo quản lý', 'Tái sử dụng',
      'Đánh giá VinaPro', 'Đánh giá VinaKey1', 'Tổng đánh giá'],
    data.hq.map((r, i) => [
      r[0], data.program[i][3] / 1e6, r[1] / 1e6, r[2] / 1e6,
      r[6] / 1e6, r[12] / 1e6, (r[6] + r[12]) / 1e6,
    ]),
    { widths: [105, 128, 128, 104, 121, 121, 112] });
  const revenue = book.worksheets.getItem('In doanh thu');
  revenue.getRange('A1').values = [[`Doanh thu kênh bán hàng · ${prefix} · Triệu đồng`]];

  printSheet(book, 'In sản lượng', `Sản lượng và hiệu quả điểm bán · ${prefix}`,
    ['Đơn vị', 'VinaPro: Số điểm', 'Điểm PSSL', 'Tỷ lệ PSSL', 'SL PTM',
      'VinaKey1: Số điểm', 'Điểm PSSL', 'Tỷ lệ PSSL', 'SL PTM', 'Tổng điểm PSSL', 'Tổng PTM'],
    data.hq.map(r => [r[0], r[4], r[7], ratio(r[7], r[4]), r[8],
      r[10], r[13], ratio(r[13], r[10]), r[14], r[7] + r[13], r[3]]),
    { widths: [104, 95, 78, 81, 64, 95, 78, 81, 64, 94, 76],
      percent: [3, 7], integer: [1, 2, 4, 5, 6, 8, 9, 10] });
}

await mkdir(directory, { recursive: true });
const calculator = Workbook.create();
addHqPrintSheets(calculator);
printSheet(calculator, 'In BSC', `Điểm BSC theo đơn vị · ${prefix}`,
  ['Đơn vị', 'HT I.8.1', 'Điểm I.8.1', 'HT I.8.2', 'Điểm I.8.2',
    'HT I.8.3', 'Điểm I.8.3', 'BSC chung'],
  data.bsc.map(r => [r[0], r[4], r[5], r[9], r[10], r[13], r[14], r[15]]),
  { widths: [110, 91, 96, 91, 96, 91, 96, 155], percent: [1, 3, 5] });
sheet(calculator, 'HQ', [hqHeaders], data.hq, { integer: Array.from({ length: 15 }, (_, i) => i + 1) });
sheet(calculator, 'BSC', [bscHeaders], data.bsc, { percent: [3, 4, 8, 9, 13], integer: [1, 2, 6, 7] });
sheet(calculator, 'Nguồn và dữ liệu thiếu', [['Nội dung', 'Trạng thái và nguồn']], data.notes,
  { widths: [170, 700] });
await (await SpreadsheetFile.exportXlsx(calculator)).save(join(directory, `HQ và BSC - ${prefix}.xlsx`));

const weekly = Workbook.create();
addHqPrintSheets(weekly);
for (const name of data.template_tabs) {
  if (name === 'KQTH Kênh PRO, KEY1') {
    const headers = data.template_headers[name].map(r => r.slice(0, 32));
    headers[0][1] = 'Kết quả PRO và KEY1';
    headers[1][1] = 'Tổng điểm PSSL PRO + KEY1';
    sheet(weekly, name, headers,
      data.hq.map(weeklyHQ), { percent: [10, 13, 15, 19, 22, 24, 27, 31] });
  } else if (name === 'CTKD DT DD và KBH') {
    const headers = data.template_headers[name].map(r => r.slice(0, 43));
    headers[4][14] = 'VINAPRO (triệu đ)';
    headers[4][15] = 'VINAKEY1 (triệu đ)';
    headers[4][16] = 'Tổng (triệu đ)';
    sheet(weekly, name, headers, data.program.map(weeklyProgram));
  } else {
    sheet(weekly, name, [['Trạng thái']], [[missing]], { widths: [180] });
  }
}
sheet(weekly, 'Nguồn và dữ liệu thiếu', [['Nội dung', 'Trạng thái và nguồn']], data.notes,
  { widths: [170, 700] });
await (await SpreadsheetFile.exportXlsx(weekly)).save(join(directory, `BC SXKD - ${prefix}.xlsx`));

const total = data.hq.at(-1);
const format = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 });
const million = value => format.format(value / 1_000_000);
const p = text => new Paragraph({ text, spacing: { after: 120 } });
const h = (text, heading = HeadingLevel.HEADING_1) => new Paragraph({ text, heading });
const children = [
  h(`Báo cáo điều hành SXKD ${prefix.toLowerCase()}`, HeadingLevel.TITLE),
  p('Bản chạy thử phần quản trị kênh bán hàng di động'),
  h('I. Dịch vụ cốt lõi tuần'),
  h('Di động', HeadingLevel.HEADING_2),
  h('Hiệu quả kênh ngoài', HeadingLevel.HEADING_3),
  p(`Doanh thu PTM theo đơn vị giao dịch IMS đạt ${million(data.program.at(-1)[3])} triệu đồng. Doanh thu đánh giá theo đơn vị quản lý điểm đạt ${million(total[6] + total[12])} triệu đồng, trong đó VinaPro ${million(total[6])} triệu đồng và VinaKey1 ${million(total[12])} triệu đồng.`),
  p(`Sản lượng PTM đạt ${format.format(total[3])} thuê bao, gồm VinaPro ${format.format(total[8])} và VinaKey1 ${format.format(total[14])}. Số điểm có phát sinh sản lượng lần lượt là ${format.format(total[7])}/${format.format(total[4])} và ${format.format(total[13])}/${format.format(total[10])}.`),
  p(`Doanh thu tái sử dụng ghi nhận được là ${million(total[2])} triệu đồng. Số liệu VinaPro đã tra TD49 trên toàn bộ tệp IMS. Nguồn tái sử dụng VinaKey1 chưa đủ để kiểm chứng.`),
  p('Thiếu dữ liệu để nhận định điều kiện chi lương, tiến độ kế hoạch CT2.1 và chênh lệch so với tuần trước. VinaKey2 và CTV liên kết sẽ bổ sung sau.'),
  p('Thiếu dữ liệu cho chiến dịch 120 ngày, BRCĐ, MyTV và các chỉ tiêu di động khác.'),
  h('II. Dịch vụ công nghệ thông tin'),
  p(missing),
  h('Nguồn và điều kiện đối chiếu'),
  ...data.notes.filter(([name]) => ['SMCS', 'BRIS', 'Đối soát TD49'].includes(name))
    .map(([name, text]) => p(`${name}. ${text}`)),
];
const doc = new Document({
  styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
  sections: [{ properties: { page: { margin: { top: 1000, bottom: 1000, left: 1150, right: 1150 } } }, children }],
});
if (mode !== '--excel-only') {
  await writeFile(join(directory, `BC Điều hành SXKD - ${prefix}.docx`), await Packer.toBuffer(doc));
}
