import * as XLSX from "xlsx";

const headers = [
  "分类号",
  "档号",
  "案卷题名",
  "卷内序号",
  "文件编号",
  "责任者",
  "文件题名",
  "文件日期",
  "页号",
  "备注",
  "保管期限",
  "总页数",
  "立卷单位",
  "开始日期",
  "结束日期",
];

export function createArchiveWorkbookFixture(): Uint8Array {
  const rows: Array<Array<string | number>> = [
    ["档案总目录测试数据"],
    headers,
    ["以下为数据"],
    [
      "5028G01",
      "5028G01-0011-842-001",
      "中核汇能高明创楷3.58904MWp分布式光伏项目 二次设备试验",
      "1",
      "GM-842-001",
      "中核汇能",
      "二次设备调试记录",
      45759,
      "1",
      "",
      "永久",
      11,
      "中核汇能高明创楷",
      "",
      "",
    ],
  ];

  for (let index = 2; index <= 11; index += 1) {
    rows.push([
      "",
      "",
      "",
      String(index),
      `GM-842-${String(index).padStart(3, "0")}`,
      "中核汇能",
      `二次设备试验文件${index}`,
      Math.min(45765, 45759 + index - 1),
      String(index),
      "",
      "",
      "",
      "",
      "",
      "",
    ]);
  }

  rows.push(
    [
      "5028G01",
      "5028G01-0011-941-001",
      "中核汇能高明创楷3.58904MWp分布式光伏项目 预装式变电站、箱式高压室进场报审，数量清单、开箱记录、装箱单、出厂质量证明、使用说明书、图纸等文件",
      "1",
      "GM-941-001",
      "中核汇能",
      "预装式变电站进场报审",
      45770,
      "页数（1478）",
      "",
      "永久",
      1527,
      "中核汇能高明创楷",
      "",
      "",
    ],
    [
      "",
      "",
      "",
      "2",
      "GM-941-002",
      "中核汇能",
      "预装式变电站图纸",
      45771,
      "页数（49）",
      "",
      "",
      "",
      "",
      "",
      "",
    ],
  );

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "总目录");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}
