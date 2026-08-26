import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = process.argv[2];
const previewDir = process.argv[3];
if (!outputPath || !previewDir) {
  throw new Error("Usage: build_workload_template.mjs <output.xlsx> <preview-dir>");
}

const workbook = Workbook.create();
const inputSheet = workbook.worksheets.add("工作量导入");
const guideSheet = workbook.worksheets.add("填写说明");

inputSheet.showGridLines = false;
inputSheet.getRange("A1:F1").values = [[
  "需求号",
  "工作描述",
  "计划开始时间",
  "计划结束时间",
  "工时/天",
  "责任人",
]];
inputSheet.getRange("A1:F1").format = {
  fill: "#1677FF",
  font: { bold: true, color: "#FFFFFF", size: 11 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#B8D4FA" },
};
inputSheet.getRange("A1:F1").format.rowHeight = 28;
inputSheet.getRange("A2:F100").format = {
  fill: "#FBFCFE",
  font: { color: "#172033", size: 10 },
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#C9D3E1" },
};
inputSheet.getRange("A2:F100").format.rowHeight = 23;
inputSheet.getRange("A2:B100").format.horizontalAlignment = "left";
inputSheet.getRange("C2:F100").format.horizontalAlignment = "center";
inputSheet.getRange("C2:D100").format.numberFormat = "yyyy-mm-dd";
inputSheet.getRange("E2:E100").format.numberFormat = "0.00";
inputSheet.getRange("A:A").format.columnWidth = 22;
inputSheet.getRange("B:B").format.columnWidth = 46;
inputSheet.getRange("C:D").format.columnWidth = 18;
inputSheet.getRange("E:E").format.columnWidth = 13;
inputSheet.getRange("F:F").format.columnWidth = 16;
inputSheet.freezePanes.freezeRows(1);
inputSheet.getRange("E2:E100").dataValidation = {
  rule: { type: "decimal", operator: "greaterThan", formula1: 0 },
};

guideSheet.showGridLines = false;
guideSheet.getRange("A1:D1").merge();
guideSheet.getRange("A1").values = [["绩效工作量导入模板 · 填写说明"]];
guideSheet.getRange("A1:D1").format = {
  fill: "#1677FF",
  font: { bold: true, color: "#FFFFFF", size: 15 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
guideSheet.getRange("A1:D1").format.rowHeight = 34;
guideSheet.getRange("A3:D3").values = [["字段", "是否必填", "示例", "填写要求"]];
guideSheet.getRange("A4:D9").values = [
  ["需求号", "是", "SR20260724045", "必须是当前计划版本中存在的需求编号"],
  ["工作描述", "是", "完成菜单管理功能", "填写本条工作的具体内容；后端算法会参考描述识别联调、数据准备等附加项"],
  ["计划开始时间", "是", "2026-08-20", "建议使用 yyyy-mm-dd；也兼容 yyyy/mm/dd 和 M.d"],
  ["计划结束时间", "是", "2026-08-22", "不得早于计划开始时间"],
  ["工时/天", "是", 2.5, "填写人天数，必须大于 0；1 天按 8 小时计算"],
  ["责任人", "是", "张三", "必须与绩效系统当前录入分组中的开发资源姓名一致"],
];
guideSheet.getRange("A3:D3").format = {
  fill: "#DCEBFF",
  font: { bold: true, color: "#134A8E" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: "#B8CBE3" },
};
guideSheet.getRange("A4:D9").format = {
  font: { color: "#172033", size: 10 },
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    insideHorizontal: { style: "thin", color: "#E1E8F0" },
    bottom: { style: "thin", color: "#B8CBE3" },
  },
};
guideSheet.getRange("A4:B9").format.horizontalAlignment = "center";
guideSheet.getRange("C4:D9").format.horizontalAlignment = "left";
guideSheet.getRange("C8").format.numberFormat = "0.00";
guideSheet.getRange("A11:D11").merge();
guideSheet.getRange("A11").values = [[
  "使用步骤：在“工作量导入”页逐行填写 → 客户端选择计划版本和录入分组 → 加载并预览 → 确认校验结果后提交。",
]];
guideSheet.getRange("A11:D11").format = {
  fill: "#ECFDF5",
  font: { bold: true, color: "#047857" },
  wrapText: true,
  verticalAlignment: "center",
};
guideSheet.getRange("A11:D11").format.rowHeight = 42;
guideSheet.getRange("A13:D13").merge();
guideSheet.getRange("A13").values = [[
  "注意：不要修改“工作量导入”页第一行表头；一个工作项占一行；模板同时适用于前端和后端分组。",
]];
guideSheet.getRange("A13:D13").format = {
  fill: "#FFF7E6",
  font: { bold: true, color: "#9A5B00" },
  wrapText: true,
  verticalAlignment: "center",
};
guideSheet.getRange("A13:D13").format.rowHeight = 42;
guideSheet.getRange("A:A").format.columnWidth = 20;
guideSheet.getRange("B:B").format.columnWidth = 13;
guideSheet.getRange("C:C").format.columnWidth = 23;
guideSheet.getRange("D:D").format.columnWidth = 64;
guideSheet.getRange("A4:D9").format.rowHeight = 38;
guideSheet.freezePanes.freezeRows(3);

const summary = await workbook.inspect({
  kind: "workbook,sheet",
  maxChars: 5000,
  tableMaxRows: 15,
  tableMaxCols: 8,
});
console.log(summary.ndjson ?? summary);
const inputRegion = await workbook.inspect({
  kind: "region",
  sheetId: "工作量导入",
  range: "A1:F6",
  maxChars: 3500,
});
console.log(inputRegion.ndjson ?? inputRegion);
const guideRegion = await workbook.inspect({
  kind: "region",
  sheetId: "填写说明",
  range: "A1:D13",
  maxChars: 6000,
});
console.log(guideRegion.ndjson ?? guideRegion);

await fs.mkdir(previewDir, { recursive: true });
for (const sheetName of ["工作量导入", "填写说明"]) {
  const preview = await workbook.render(
    sheetName === "工作量导入"
      ? { sheetName, range: "A1:F18", scale: 1.5, format: "png" }
      : { sheetName, autoCrop: "all", scale: 1.5, format: "png" },
  );
  const filename = sheetName === "工作量导入" ? "template-input.png" : "template-guide.png";
  await fs.writeFile(`${previewDir}/${filename}`, new Uint8Array(await preview.arrayBuffer()));
}

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 4000,
});
console.log(formulaErrors.ndjson ?? formulaErrors);

await fs.mkdir(outputPath.substring(0, outputPath.lastIndexOf("/")), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
