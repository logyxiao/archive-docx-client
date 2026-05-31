import { generateProcessDocs } from './src/lib/processDocs.ts';
import fs from 'fs';

const processRecord = {
  categoryCode: "8312",
  archiveCode: "5028G01-0011-8312-001",
  fullTitle: "光伏区安装部分",
  projectName: "中核汇能遂溪国恒坤达3.84MW屋顶分布式光伏项目",
  volumeTitle: "光伏方阵安装",
  owner: "中核华辰建筑工程有限公司",
  filingUnit: "中核华辰建筑工程有限公司",
  retentionPeriod: "30年",
  startDate: "20241208",
  endDate: "20250420",
  dateRange: "20241208-20250420",
  totalPages: 94,
  drawingPages: 0,
  textPages: 94,
  items: [
    {
      sequence: "1",
      fileCode: "5028G01-SG-ZHHC-TEST-001",
      owner: "中核华辰建筑工程有限公司",
      title: "高明分布式项目 并网点光伏发电系统单位工程开工报审",
      fileDate: "20250410",
      pageNo: "1",
      note: ""
    },
    {
      sequence: "2",
      fileCode: "5028G01-SG-ZHHC-TEST-002",
      owner: "中核华辰建筑工程有限公司",
      title: "高明分布式项目 光伏方阵安装子单位工程质量报验申请及验收记录",
      fileDate: "20250410",
      pageNo: "2",
      note: ""
    },
    {
      sequence: "3",
      fileCode: "5028G01-SG-ZHHC-TEST-003",
      owner: "中核华辰建筑工程有限公司",
      title: "高明分布式项目 子方阵场电气安装分部工程质量报验申请及验收记录",
      fileDate: "20250410",
      pageNo: "3",
      note: ""
    },
    {
      sequence: "20",
      fileCode: "5028G01-SG-ZHHC-TEST-020",
      owner: "中核华辰建筑工程有限公司",
      title: "高明分布式项目 #2光伏升压变电缆防火阻燃施工分项工程质量报验申请及验收记录",
      fileDate: "20250410",
      pageNo: "20",
      note: ""
    }
  ]
};

const record8312002 = {
  ...processRecord,
  archiveCode: "5028G01-0011-8312-002",
  fullTitle: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目 光伏变电系统子单位工程质量报审，分部、分项及检验批质量验收记录文件",
  items: [
    {
      ...processRecord.items[2],
      sequence: "3",
      fileCode: "5028G01-SG-ZHHC-02-01-001",
      title: "高明分布式项目 子方阵场电气安装分部工程质量报验申请及验收记录",
    },
    {
      ...processRecord.items[3],
      sequence: "20",
      fileCode: "5028G01-SG-ZHHC-02-02-03-002",
      title: "高明分布式项目 #2光伏升压变电缆防火阻燃施工分项工程质量报验申请及验收记录",
    },
  ],
};

globalThis.fetch = async (input) => {
  const path = String(input).replace(/^https?:\/\/[^/]+/, "");
  return new Response(fs.readFileSync(`public${decodeURIComponent(path)}`));
};

async function main() {
  const result = await generateProcessDocs(
    [record8312002],
    {
      selectedCodes: [record8312002.archiveCode],
      outputDir: "/tmp/archive-output",
    },
    async (path) => {
      console.log('Generating:', path);
    },
  );
  
  console.log('Result files:', result.files.length);
  console.log('Result errors:', result.errors);
}

main().catch(console.error);
