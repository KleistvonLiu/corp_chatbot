import { execFile } from "node:child_process";
import { promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import { parseWorkflowWorkbook } from "../server/lib/parsers";

type CanonicalEntryType = "流程" | "联系人" | "供应商" | "参考" | "系统链接";

interface CanonicalRow {
  rowNumber: number;
  category: string;
  entryType: CanonicalEntryType;
  title: string;
  relatedForm?: string;
  contacts?: string;
  body: string;
  url?: string;
  keywords: string[];
  imageLinks?: string[];
}

interface ImageAttachment {
  rowNumber: number;
  pageNumber: number;
  fileName: string;
  relativePath: string;
  attachmentType: "png";
  description: string;
}

const outputDir = "/home/kleist/Downloads/corp-eng-new-staff-guide-20260401-canonical";
const zipPath = "/home/kleist/Downloads/corp-eng-new-staff-guide-20260401-canonical.zip";
const sourcePdfPath = "/home/kleist/Downloads/Corp. Eng New Staff Guide Book-20260401.pdf";
const execFileAsync = promisify(execFile);

function joinLines(...parts: Array<string | undefined>) {
  return parts
    .flatMap((part) => (part == null ? [] : part.split("\n")))
    .map((line) => line.trimEnd())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function pad2(value: number) {
  return String(value).padStart(2, "0");
}

const imagePagesByRow = new Map<number, number[]>([
  [1, [4]],
  [2, [5]],
  [3, [6]],
  [4, [7]],
  [5, [8]],
  [6, [9]],
  [7, [9]],
  [8, [10]],
  [9, [10]],
  [10, [11]],
  [11, [11]],
  [12, [12]],
  [13, [13]],
  [14, [13]],
  [15, [13]],
  [16, [14]],
  [17, [15]],
  [18, [16]],
  [19, [18]],
  [20, [19]],
  [21, [20]],
  [22, [21]],
  [23, [22]],
  [24, [23]],
  [25, [24]],
  [26, [25]],
  [27, [26]],
  [28, [27]],
  [29, [28]],
  [30, [29]],
  [31, [30]],
  [32, [31]],
  [33, [32]],
  [34, [33]],
  [35, [34]],
  [36, [35]],
  [37, [36]],
  [38, [37, 38]],
  [39, [39, 40]],
  [40, [41]],
  [41, [42]],
  [42, [43]],
  [43, [44]],
  [44, [45]],
  [45, [46]],
  [46, [47]],
  [47, [48]],
  [48, [49]],
  [49, [50]],
  [50, [51]],
  [51, [52]],
  [52, [53]],
  [53, [54]],
  [54, [55]],
  [55, [56]],
  [56, [57]],
  [57, [58]],
  [58, [59]],
  [59, [60]],
  [60, [61]],
  [61, [62]],
  [62, [63]],
  [63, [64]],
  [64, [65]],
  [65, [66]],
  [66, [67]],
  [67, [68]],
  [68, [69]],
  [69, [70]],
  [70, [71]],
  [71, [72, 73]],
  [72, [74]]
]);

const rows: CanonicalRow[] = [
  {
    rowNumber: 1,
    category: "组织与导览",
    entryType: "参考",
    title: "新同事共享册介绍",
    body: joinLines(
      "欢迎加入集团工程部大家庭。",
      "此信息共享册用于帮助新同事更快了解德昌电机相关支持部门的政策、规定、流程和集团工程部的基本情况，便于尽快融入团队。",
      "本文件仅供集团工程部内部使用，请勿向外传阅。"
    ),
    keywords: ["新同事共享册", "入职介绍", "新人手册", "Corp. Eng", "Guide Book"]
  },
  {
    rowNumber: 2,
    category: "组织与导览",
    entryType: "参考",
    title: "P200厂区平面图",
    body: joinLines(
      "P200 区总平面图可辨识区域如下：",
      "1. 主体建筑区包括 B1、B2、B3、B3A、B8、B9、B10。",
      "2. 生活/配套区包括饭堂以及 D1、D2、D3、D4、D5。",
      "3. 出入口和外部参照包括北大门、南大门、门卫位置。",
      "4. 图中还可见创世纪、新杰物流等外围参照区域。",
      "5. 如需现场定位，可优先用 B 区/D 区楼号、饭堂、南北门做方向锚点。"
    ),
    keywords: ["P200", "厂区平面图", "B1", "B2", "B3", "B8", "B9", "B10", "饭堂", "南大门", "北大门"]
  },
  {
    rowNumber: 3,
    category: "组织与导览",
    entryType: "参考",
    title: "江门厂区平面图",
    body: joinLines(
      "江门厂区平面图可辨识区域如下：",
      "1. 出入口区域包括正门、后门、3号门（供应商进出）以及接待处/安防部。",
      "2. 住宿与生活区包括 H1 旅馆、D2 宿舍楼、D1/1F 宿舍服务处、HR 办公室、医务室。",
      "3. 实验室与办公区包括 D3/1F&2F Corp 实验室、APG、IPQ 实验室。",
      "4. B2/3F 为半导体实验室，B2/6F 为材料分析实验室。",
      "5. B1/1F 为材料分析实验室，B1/2F 为 EMC 实验室与检定部，B1/3F 为 IT 办公室。",
      "6. 图中还标出后门停车场、前门停车场等位置。"
    ),
    keywords: ["江门", "厂区平面图", "3号门", "正门", "后门", "H1旅馆", "D2宿舍楼", "EMC实验室", "IT办公室"]
  },
  {
    rowNumber: 4,
    category: "组织与导览",
    entryType: "参考",
    title: "集团工程部各组相关介绍",
    body: joinLines(
      "集团工程部主要分工如下：",
      "1. Allan Kwan：工程运营、流程、财务、规格、行政支持。",
      "2. Yu Chen：设计标准、生产支持、马达改善、DFM/DFA、EMC 测试、马达/齿轮设计仿真、NVH 测试与分析。",
      "3. Ning Sun：新产品机械研发。",
      "4. Janet Fang：风机和水泵开发、流体和热仿真。",
      "5. Xia Chen、Jun Fu、Stanley Wu、PingHua Tang：新产品电子研发。",
      "6. XiaoHong Zhou：项目管理、样板制作与机加工。",
      "7. Calvin Yuen：材料分析、材料工程、JEEP 电子封装中心。",
      "8. Downing Tang：德昌样板部与实验室。"
    ),
    keywords: ["集团工程部", "各组介绍", "Allan Kwan", "Yu Chen", "Ning Sun", "Janet Fang", "Calvin Yuen", "Downing Tang"]
  },
  {
    rowNumber: 5,
    category: "组织与导览",
    entryType: "参考",
    title: "集团工程部支持组工作介绍",
    body: joinLines(
      "集团工程部支持组分工如下：",
      "1. Allan Kwan：工程运营、流程、财务、预算控制、财务支持。",
      "2. Helen Wang：行政支持、考勤管理、文具管理、工具与仪器检定、资产管理、软件管理、办公室管理、装修、报销、马达报废、6S&EHS&消防管理、活动组织（年会、Tech day）等。",
      "3. ManHin Wong：工程规格。",
      "4. Li Sang：PDMS IV 管理、N 盘文件服务器管理、Solid Edge/CAD/CAM 支持、Team Center 支持、其他工程系统支持。",
      "5. Susie Huang：BU HR。"
    ),
    keywords: ["支持组", "Allan Kwan", "Helen Wang", "ManHin Wong", "Li Sang", "Susie Huang", "PDMS IV", "Team Center"]
  },
  {
    rowNumber: 6,
    category: "行政支持",
    entryType: "联系人",
    title: "JiaLian Li（李家连）",
    contacts: "JiaLian Li；李家连",
    body: joinLines(
      "联系方式：Ext 3325；手机 13714085757。",
      "主要职责：内部采购单 PR、考勤管理、资产管理、工具仪器检定（Shajing）、集团工程部门禁权限管理、办公室 6S 检查与问题跟进、申请德昌电机内部短号群。",
      "详细事项：",
      "1. 管理和发放办公易耗品（文具、饮用水、纸巾等），盘点库存并及时购买；管理储物柜和移动柜；根据需求开 PR 采购研发物料和测试仪器并跟进送货。",
      "2. 每天处理考勤异常通知、资料收集和月结工时核对；集中递交需要 HR 加签的文件。",
      "3. 做资产盘点，收集新资产型号/序列号并提交系统，跟进资产转移和报废登记。",
      "4. 根据检定清单收集待检仪器并送检；做新仪器/工具注册登记；跟进检定进度、收回已检仪器、处理异常仪器维修/报废，以及离职同事名下仪器工具的转移与报失。",
      "休假代理人：李捷、宋丹丹、陈燕妮。"
    ),
    keywords: ["JiaLian Li", "李家连", "PR", "考勤", "资产", "仪器检定", "门禁", "6S"]
  },
  {
    rowNumber: 7,
    category: "行政支持",
    entryType: "联系人",
    title: "YanNi Chen（陈燕妮）",
    contacts: "YanNi Chen；陈燕妮",
    body: joinLines(
      "联系方式：Ext 3325；手机 18273726408。",
      "主要职责：办公室管理与设备维护、实验室维修跟进、集团工程部会议室预订和管理、VC 预订、报销（EAR）单递交、小车申请、穿梭巴士安排、货车与线路巴、小车位申请、集团工程部 UPR 申请（邮箱/电脑权限）、办公室 6S 检查、活动策划及其他行政支持。",
      "详细事项：",
      "1. 留意办公室异味、绿植、空调温度变化并及时处理；跟进打印机、空调、灯管、摄像头等维修；管理药品箱。",
      "2. 发出并跟进部门维修单，包括试验室内插座、排气扇、风扇、空调、水电等维修。",
      "3. 会议室与 VC 预订需至少提前一天，并在预订时说明设备支持需求。",
      "4. 协助出差等报销申请，建议出差两个月内完成报销流程，准备 TRF、登机牌、发票等材料。",
      "5. 小车申请需提前一天 16:30 前交单。",
      "6. 日常还负责部门信件领取、物品/文件邮递、笔记本相机进出申请、工资卡/社保卡/住房公积金卡发放和加盖公章业务。",
      "休假代理人：李家连、宋丹丹、李捷。"
    ),
    keywords: ["YanNi Chen", "陈燕妮", "EAR", "UPR", "会议室", "小车申请", "穿梭巴士", "货车", "邮递"]
  },
  {
    rowNumber: 8,
    category: "行政支持",
    entryType: "联系人",
    title: "Jie Li（李捷）",
    contacts: "Jie Li；李捷",
    body: joinLines(
      "联系方式：Ext 3325；手机 18566232786。",
      "主要职责：雇员出差机票及酒店预订、ITPM 申请与软件/电脑申购管理、集团工程部 EHS/6S 监督执行、危废仓管理及危废报废安排、大型机械设备搬运流程监督、TISAX 审核跟进、大型活动支持与执行。",
      "详细事项：",
      "1. 机票/酒店：需准备 TRF 订机票；预订香港酒店建议每周一前告知。",
      "2. ITPM：协助申请 IT 权限、办公电脑/工业电脑、IT 设备采购 PR 单，跟进软件及相关配备、电脑/软件转移和台账更新。",
      "3. EHS/6S：每月 28 号左右召集安全主任开会，组织消防演练、消防培训、更新消防架构图和疏散图，组织 6S 检查及日常巡查。",
      "4. 装修/JR：跟进办公室及试验室施工工程请求单、报价、施工开工会、质量与验收。",
      "休假代理人：宋丹丹、陈燕妮、李家连。"
    ),
    keywords: ["Jie Li", "李捷", "ITPM", "EHS", "6S", "TISAX", "机票", "酒店", "危废"]
  },
  {
    rowNumber: 9,
    category: "行政支持",
    entryType: "联系人",
    title: "DanDan Song（宋丹丹）",
    contacts: "DanDan Song；宋丹丹",
    body: joinLines(
      "联系方式：手机 15889566312；分机 616312。",
      "主要职责：现金采购和报销管理及特殊物品采购（Non-PO）、活动统筹及支持、跨部门会议、办公室装修/JR 单跟进、其他事务。",
      "详细事项：",
      "1. 负责办公室及试验室施工工程请求单的发出，跟进图纸、报价、施工进度、工程质量与验收，并协调物品搬运和外部搬运报价。",
      "2. 负责现金采购、报销管理和 Non-PO 特殊采购。",
      "3. 参与并推动年度聚餐、团建、聚餐津贴申领等活动。",
      "4. 参与 6S&FAST Meeting、公司 6S 巡查、饭堂膳食会议、爱心基金申请等事项。",
      "5. 跟进非资产报废申请、数据报表制作和跨部门协调。",
      "休假代理人：李捷、陈燕妮、李家连。"
    ),
    keywords: ["DanDan Song", "宋丹丹", "Non-PO", "现金采购", "报销", "活动统筹", "JR单"]
  },
  {
    rowNumber: 10,
    category: "行政支持",
    entryType: "联系人",
    title: "AnHui Hu（胡安慧）",
    contacts: "AnHui Hu；胡安慧",
    body: joinLines(
      "联系方式：手机 13534274780。",
      "主要职责：楼面（含 R&D）日常收货、日常报废马达及非资产物品处理、6S/安全/消防工作、节假日福利品发放、桶装水管理、临时行政支持。",
      "详细事项：",
      "1. 每天到 7 座易耗品仓收集并分发部门货物，负责签收、运送、退货及异常问题跟进。",
      "2. 负责 Corp. Eng 马达报废和非资产物品报废，跟进报废单开出、签批和送达报废厂。",
      "3. 做部门每日 6S/消防巡查、设施设备检查和消防/EHS 看板更新，配合消防演练与宣传。",
      "4. 负责公司福利品发放以及桶装水数量确认、送水位置协调和漏送问题跟进。",
      "休假代理人：郭威。"
    ),
    keywords: ["AnHui Hu", "胡安慧", "收货", "报废马达", "非资产报废", "消防", "福利品", "桶装水"]
  },
  {
    rowNumber: 11,
    category: "行政支持",
    entryType: "联系人",
    title: "Wei Guo（郭威）",
    contacts: "Wei Guo；郭威",
    body: joinLines(
      "联系方式：分机 817-2047；手机 18786500219。",
      "主要职责：江门办公室行政事务、江门办公室管理、江门各试验室日常维修工作跟进、KM 工具仪器管理、日常报废马达处理。",
      "详细事项：",
      "1. 管理和发放办公易耗品与福利品；每天下午送交小车申请单到车队并跟进；负责物料收发与快递寄送。",
      "2. 跟进江门办公室异味、绿植、空调、门禁、摄像头等异常；跟进办公室物品、TISAX 审核、江门门禁权限申请和 C1-4F 会议室管理。",
      "3. 负责 C1-4F 电子测试区、D3-1F&2F Corp. Eng 试验室、B1-2F EMC 试验室的装修/维修单、报价、施工和验收跟进。",
      "4. 根据检定清单收集、登记、送检、回收并分发 KM 工具仪器，处理异常仪器后续事项。",
      "5. 协助做好报废马达拆解与送报废工作。",
      "休假代理人：陈燕妮、李家连、宋丹丹、李捷；KM 工具仪器管理事项的主要代理人为胡安慧。"
    ),
    keywords: ["Wei Guo", "郭威", "江门行政", "门禁", "TISAX", "EMC实验室", "KM工具仪器", "报废马达"]
  },
  {
    rowNumber: 12,
    category: "行政支持",
    entryType: "参考",
    title: "公司内部电话区号查询信息",
    body: joinLines(
      "可通过 JE OnNet Site List 查询公司内部电话区号信息。",
      "页面截图显示 Last Update: 12 Oct 2018。",
      "如需最新区号信息，建议以公司内网实际页面为准。"
    ),
    keywords: ["电话区号", "JE OnNet Site List", "内线", "分机", "区号查询"]
  },
  {
    rowNumber: 13,
    category: "行政支持",
    entryType: "系统链接",
    title: "沙井员工联系方式查询",
    body: joinLines(
      "沙井员工联系方式可通过集团内网 SharePoint 通讯录查询。",
      "当前手册给出的入口是 301 Telephone Directory（Jun.2023 版本）。",
      "如链接跳转失败，可在 JE In Motion 站点中搜索电话通讯录。 "
    ),
    url: "https://jehl.sharepoint.com/:x:/r/sites/JEInMotion/_layouts/15/doc2.aspx?sourcedoc=%7B2288E4FE-56EC-4EE7-B3BF-860A95BDBB69%7D&file=301-Telephone-Directory-Jun.2023-.xls&action=default&mobileredirect=true&wdLOR=cC2DA9994-0641-41BB-AFBB-BE43F44D6D21",
    keywords: ["沙井员工", "联系方式", "电话通讯录", "301 Telephone Directory", "SharePoint"]
  },
  {
    rowNumber: 14,
    category: "行政支持",
    entryType: "系统链接",
    title: "香港员工联系方式查询",
    body: joinLines(
      "香港员工联系方式可通过集团内网 SharePoint 通讯录查询。",
      "当前手册给出的入口文件名为 SP---P12-telephone-list-Sep-2025.xls。"
    ),
    url: "https://jehl.sharepoint.com/:x:/r/sites/JEInMotion/_layouts/15/Doc.aspx?sourcedoc=%7B1644CF10-0D27-42A2-8EAC-46D201356ADC%7D&file=SP---P12-telephone-list-Sep-2025.xls&action=default&mobileredirect=true&wdLOR=c7D6B8306-1E35-4704-B50E-0DBEC1D317FD",
    keywords: ["香港员工", "联系方式", "电话通讯录", "P12 telephone list", "SharePoint"]
  },
  {
    rowNumber: 15,
    category: "行政支持",
    entryType: "系统链接",
    title: "江门员工联系方式查询",
    body: joinLines(
      "江门员工联系方式可通过集团内网 SharePoint 通讯录查询。",
      "当前手册给出的入口文件名为 KM-Telephone-Directory--June-2024-.xlsx。"
    ),
    url: "https://jehl.sharepoint.com/:x:/r/sites/JEInMotion/_layouts/15/Doc.aspx?sourcedoc=%7B2306985E-00A4-468B-B451-6CEF581AAF9C%7D&file=KM-Telephone-Directory--June-2024-.xlsx&action=default&mobileredirect=true&wdLOR=c57C0D5D9-BD3C-458B-A0CF-925873FFCD7A",
    keywords: ["江门员工", "联系方式", "电话通讯录", "KM Telephone Directory", "SharePoint"]
  },
  {
    rowNumber: 16,
    category: "组织与导览",
    entryType: "参考",
    title: "职员生日会",
    body: joinLines(
      "职员生日会是公司福利政策之一，由部门和 BUHR 统一安排组织。",
      "一般每季度举行一次。",
      "活动通常包含甜点零食、蛋糕和团队游戏。"
    ),
    keywords: ["生日会", "福利政策", "BUHR", "季度活动"]
  },
  {
    rowNumber: 17,
    category: "组织与导览",
    entryType: "参考",
    title: "年度聚餐（晚会）",
    body: joinLines(
      "年度聚餐/晚会是公司福利政策之一，由部门统一安排组织。",
      "一般安排在中国农历新年前后举行。",
      "活动通常包含晚餐、各组别节目、游戏和抽奖。"
    ),
    keywords: ["年度聚餐", "晚会", "年饭", "福利政策", "抽奖"]
  },
  {
    rowNumber: 18,
    category: "行政支持",
    entryType: "流程",
    title: "爱心互助基金",
    relatedForm: "爱心互助基金申请表",
    contacts: "DanDan Song",
    body: joinLines(
      "德昌电机“爱心互助基金”于 2009-01-01 成立，由企业工会发起，属于雇员互助性质的共济基金。",
      "加入爱心互助基金的雇员，在职期间罹患疾病或非因公受伤时可以申请。",
      "表单参考：Forms\\爱心互助基金救助申请表2025最新版.doc。",
      "申请时需提供给 DanDan Song 的资料包括：《爱心互助基金申请表》、住院证明、出院小结、发票原件、住院费用结算清单原件、厂牌及身份证复印件；门诊治疗如符合条件，还需提供门诊病历和费用清单。",
      "手册页面另附“救助标准”图示，具体额度建议以行政组和基金当期执行标准为准。"
    ),
    keywords: ["爱心互助基金", "互助基金", "基金申请", "DanDan Song", "住院证明", "门诊病历"]
  },
  {
    rowNumber: 19,
    category: "IT",
    entryType: "流程",
    title: "电脑权限申请流程",
    relatedForm: "UPR单",
    contacts: "YanNi Chen",
    body: joinLines(
      "新同事入职前通常已开通基本电脑权限，例如电脑登录账号和邮箱账号，可找 Team Leader 或 YanNi Chen 领取密码。",
      "如需其他权限（如 PLM、PDMS IV、ERP 等），可发邮件给 YanNi Chen 开 UPR 单，邮件中写明姓名、工号、需要开通的权限名称及用途。",
      "UPR 单需经总监和 IT 部门审核批准，再由电脑部开通权限。"
    ),
    keywords: ["电脑权限", "UPR", "PLM", "PDMS IV", "ERP", "YanNi Chen"]
  },
  {
    rowNumber: 20,
    category: "IT",
    entryType: "流程",
    title: "电脑/需联网硬件设备采购",
    relatedForm: "ITPM申请表；PR单",
    contacts: "Jie Li",
    body: joinLines(
      "1. 发邮件给 Jie Li 申请提交 ITPM，并在 ITPM 申请表中注明电脑类别/硬件设备信息和所属 Owner 信息。",
      "2. 若所申请电脑属于 JE 标准电脑，无需提供报价；如是非标电脑，则需提供报价表。",
      "3. 提交 ITPM 申请表时，需要一并提交 Team Leader 和 Cost Center Head 的批准邮件。",
      "4. 待 IT 部批准 ITPM 后，才可开 PR 购买。",
      "5. 电脑设备如果转移使用人或转移部门，也需要通过 Jie Li 开 ITPM 并更新信息。"
    ),
    keywords: ["电脑采购", "联网硬件", "ITPM", "PR", "Jie Li", "Owner"]
  },
  {
    rowNumber: 21,
    category: "IT",
    entryType: "流程",
    title: "软件购买、安装",
    relatedForm: "ITPM申请表；PR单；SD单",
    contacts: "Jie Li",
    body: joinLines(
      "软件购买：",
      "1. 发邮件给 Jie Li 申请提交 ITPM，申请表中注明软件名称、版本、数量等信息。",
      "2. 除 ITPM 申请表外，还需提供报价表，以及 Team Leader 和 Cost Center Head 的批准邮件。",
      "3. 待 IT 部批准 ITPM 后，再开 PR 单购买。",
      "4. 待采购通知软件 license 到货后，再开 SD 单至 IT 登记 license 并安装。",
      "",
      "软件安装：",
      "1. 发邮件给 Jie Li 申请提交 ITPM，申请表中注明软件对应的 PR/PO、名称、版本等信息。",
      "2. 如为免费可商用软件，除 ITPM 外还需提供官方免费声明、下载链接以及对应硬件设备信息。",
      "3. 提交 ITPM 时同样需要 Team Leader 和 Cost Center Head 的批准邮件。",
      "4. 待 IT 部批准 ITPM 后，再开 SD 单安装。"
    ),
    keywords: ["软件购买", "软件安装", "ITPM", "SD单", "PR", "PO", "license", "Jie Li"]
  },
  {
    rowNumber: 22,
    category: "IT",
    entryType: "流程",
    title: "电脑问题求助方式",
    relatedForm: "SD单",
    body: joinLines(
      "电脑使用中遇到自己无法解决的问题，可按以下方式求助：",
      "1. 开 SD 单，在 CA Service Desk Manager – Home 录入信息。",
      "2. 电话联系 IT 工程师：P200-3B-2F 郭兴华 13682541730；KM-C1-4F 黄家辉 15917376164。",
      "3. 发邮件描述问题：APAC Service Desk/CHN/JEHL。",
      "4. 热线：P200 Ext 2222 / Tel 0755-29900945；KM Tel 0750-3202197。",
      "5. 直接到 P200 区 10 座 2 楼或 KM 区 B1 栋 3 楼电脑部现场求助。"
    ),
    keywords: ["电脑问题", "SD单", "Service Desk", "郭兴华", "黄家辉", "2222", "APAC Service Desk"]
  },
  {
    rowNumber: 23,
    category: "IT",
    entryType: "流程",
    title: "打印机管理",
    contacts: "JiaLian Li；Wei Guo",
    body: joinLines(
      "由于 TISAX information security 信息安全要求，公司使用理光刷卡安全打印机。",
      "目前 3 座 2 楼有 4 台黑白和 2 台彩色打印机（A4/A3/A2/A1 纸张），KM C1-4F 有 1 台彩色打印机（A4/A3 纸张）。",
      "打印机安装流程：",
      "1. 管理员注册新用户账号：P200 管理员为 JiaLian Li，KM 管理员为 Wei Guo。",
      "2. 用户携带厂牌到打印机处刷卡完成注册。",
      "3. 用户自行在电脑安装打印机：P200 使用 //eprinter on carp020008；KM 使用 //eprinter on carp030011。",
      "4. 如无法自行安装，请找电脑部同事安装驱动。"
    ),
    keywords: ["打印机", "理光", "刷卡打印", "eprinter", "TISAX", "JiaLian Li", "Wei Guo"]
  },
  {
    rowNumber: 24,
    category: "行政支持",
    entryType: "流程",
    title: "工衣管理",
    relatedForm: "企业微信审批",
    body: joinLines(
      "新雇员入职当天可在 HR 工衣管理员曾晓燕（Ext 2269）处领取冬、夏各 2 件工衣，并登记到个人名下；离职时需要将工衣交回 HR 离职组。",
      "上班期间必须穿工衣并佩戴厂牌。",
      "工衣破损或有污渍时，可在企业微信 App 的“工作台-审批”中申请更换。",
      "手册页面附有“工衣更换标准”图示，执行口径建议以 HR 当期规则为准。"
    ),
    keywords: ["工衣", "曾晓燕", "企业微信审批", "工衣更换", "厂牌"]
  },
  {
    rowNumber: 25,
    category: "行政支持",
    entryType: "流程",
    title: "文具、柜子管理发放",
    contacts: "JiaLian Li；YanNi Chen；DanDan Song；Jie Li",
    body: joinLines(
      "1. 新入职员工可根据工作需要申请移动柜或文件柜，相关分配信息由 JiaLian Li 统一登记。",
      "2. 柜内仅可存放个人办公用品、工衣等私人物品，严禁存放易燃易爆、危险品、贵重物品及违规物品。",
      "3. 离职时需将领取物品、清理好的柜子以及钥匙退还给 JiaLian Li。",
      "4. 文具领用按工号或姓名登记即可。",
      "5. 文具发放人员包括家连、燕妮、丹丹、李捷。"
    ),
    keywords: ["文具", "柜子", "移动柜", "文件柜", "JiaLian Li", "发放"]
  },
  {
    rowNumber: 26,
    category: "人事/考勤",
    entryType: "流程",
    title: "电子工资单的查看方法",
    body: joinLines(
      "大陆同事的工资通常在每月 9 号至 12 号转入工资卡。",
      "电子工资条可直接在企业微信中查询。",
      "手册页面主要提供了操作截图，具体界面路径建议以企业微信当前版本为准。"
    ),
    keywords: ["电子工资单", "工资条", "企业微信", "工资卡"]
  },
  {
    rowNumber: 27,
    category: "人事/考勤",
    entryType: "流程",
    title: "工作时间",
    contacts: "JiaLian Li",
    body: joinLines(
      "集团工程部一般实行“五天八小时”工作制。",
      "常规上班时间为 08:30-12:00、13:00-17:30，中午休息 1 小时。",
      "如需调整班次，可申请转班，并提前与 JiaLian Li（Ext 3325）确认目标班次是否可用。",
      "示例：9:30 班次时段为 09:30-12:15、12:15-18:30。"
    ),
    keywords: ["工作时间", "五天八小时", "转班", "JiaLian Li", "9:30班次"]
  },
  {
    rowNumber: 28,
    category: "人事/考勤",
    entryType: "流程",
    title: "刷卡",
    relatedForm: "跨楼栋刷卡申请表",
    body: joinLines(
      "1. 公司采用 IC 卡考勤系统，每天需刷 4 次上下班卡，特殊情况除外。",
      "2. 每次刷卡间隔 11 分钟后再刷卡才算有效卡；间隔不足 11 分钟为无效卡。",
      "3. 雇员原则上应在本部门所在楼栋楼层刷考勤卡；长期乘坐公司班车的雇员可在外勤卡钟刷卡。",
      "4. 如因特殊需要必须跨楼栋刷卡，需要递交《跨楼栋刷卡申请表》。"
    ),
    keywords: ["刷卡", "IC卡", "考勤卡", "跨楼栋刷卡申请表", "外勤卡钟"]
  },
  {
    rowNumber: 29,
    category: "人事/考勤",
    entryType: "流程",
    title: "未录卡补卡",
    relatedForm: "301区雇员考勤动态表；补卡申请表",
    contacts: "JiaLian Li",
    body: joinLines(
      "1. 雇员按时到岗但未刷卡时，应填写《301区雇员考勤动态表》申报补卡，经 Team Leader 签批后交 JiaLian Li 处理，否则按旷工计算。",
      "2. 在一个自然月内，因个人原因未录卡（厂证损坏不能打卡及公务除外）最多可补三次，超过三次按旷工处理。",
      "3. 新入职直接与非直接员工：第二天第一个上班考勤卡默认有效，无需补卡。",
      "4. 新入职职员：第一天全天及第二天第一个上班考勤卡默认上下班卡，无需补卡。",
      "5. 离职当天需完成 3 次打卡（上午上班卡、上午下班卡、下午上班卡），方可享受当天全天工资。",
      "6. 厂证损坏或遗失时，应先在企业微信 App 中申请，经 HR 审批后到南门招聘中心厂证办理处（Ext 3825）补办；期间未录卡需补卡，并由厂证办理负责人签字确认。",
      "7. 班车迟到补卡，可在企业微信“工作台-审批”中申请。",
      "8. 如尚未加入企业微信，可找 JiaLian Li 协助。"
    ),
    keywords: ["未录卡", "补卡", "考勤动态表", "JiaLian Li", "厂证", "企业微信"]
  },
  {
    rowNumber: 30,
    category: "人事/考勤",
    entryType: "流程",
    title: "出差考勤",
    relatedForm: "补卡申请表",
    body: joinLines(
      "KM、深圳等短途出差，直接填写《补卡申请表》即可。",
      "外地出差补卡需附凭证：",
      "1. 飞机出行：附携程商旅订票成功截图或机票图片（打印件）。",
      "2. 火车出行：提供火车票（打印件）。",
      "3. 打车出行：提供行程单或发票（打印件）。",
      "4. 如出差期间跨周休日或法定假日，需按公司考勤规则执行。"
    ),
    keywords: ["出差考勤", "补卡申请表", "携程商旅", "火车票", "行程单", "发票"]
  },
  {
    rowNumber: 31,
    category: "人事/考勤",
    entryType: "参考",
    title: "迟到/早退",
    body: joinLines(
      "考勤扣工时（按缺勤工时计）规则如下：",
      "1. 1 分钟内：不计缺勤工时。",
      "2. 2-10 分钟：缺勤 0.25 小时（即 0.03 天）。",
      "3. 11-30 分钟：缺勤 0.5 小时（即 0.06 天）。",
      "4. 31-60 分钟：缺勤 1 小时（即 0.125 天）。",
      "5. 61-90 分钟：缺勤 1.5 小时（即 0.188 天）。",
      "6. 91-120 分钟：缺勤 2 小时（即 0.25 天）。",
      "7. 120-240 分钟：缺勤 4 小时（即 0.5 天）。",
      "8. 超过 240 分钟：缺勤 8 小时（即 1 天）。"
    ),
    keywords: ["迟到", "早退", "扣工时", "缺勤", "考勤计算"]
  },
  {
    rowNumber: 32,
    category: "人事/考勤",
    entryType: "流程",
    title: "加班",
    relatedForm: "加班申请表；法定假日加班申请表",
    body: joinLines(
      "1. 加班需提前提交《加班申请表》。",
      "2. 工作日加班：8 小时之外延长工作时间视为加班，按 1.5 倍加班工资计算。DL、IDL 及部分有加班费的 Staff 当月计发加班费；无加班费的 Staff 不计加班费。",
      "3. 周休日加班：按 2 倍加班工资计算。有加班费的 Staff 可计加班费，也可按 4 小时为最小单位折算调休；无加班费的 Staff 可安排调休。",
      "4. 法定假日加班：按 3 倍加班工资计算，需提前一周提交由 GM 或 SVP 批准、且人力资源部总监或以上人员批准的《法定假日加班申请表》。"
    ),
    keywords: ["加班", "加班申请表", "法定假日加班", "调休", "DL", "IDL"]
  },
  {
    rowNumber: 33,
    category: "人事/考勤",
    entryType: "流程",
    title: "调休",
    relatedForm: "雇员考勤动态表；加班申请表；请假申请表",
    body: joinLines(
      "调休即周末加班与正常班调换，原则上遵循“先加班后调休”。",
      "1. 加班当月调休：提交经 Team Leader 审批后的《雇员考勤动态表》。",
      "2. 跨月调休：有效期为 3 个月，需提交《加班申请表》和《请假申请表》。"
    ),
    keywords: ["调休", "考勤动态表", "加班申请表", "请假申请表", "跨月调休"]
  },
  {
    rowNumber: 34,
    category: "人事/考勤",
    entryType: "流程",
    title: "年假、轮休假",
    relatedForm: "请假申请表；职员年假单",
    contacts: "JiaLian Li",
    body: joinLines(
      "1. 职员年假：半天为最小请假单位，需填写《请假申请表》与《职员年假单》，经 Team Leader 审批后交 JiaLian Li。",
      "2. 可休年假天数 = 当年度已在德昌服务月数 / 12 × 当年年假天数。",
      "3. 例：若员工全年年假为 14 天，则 4 月份可休 4.5 天（4 ÷ 12 × 14 = 4.6，按 4.5 天执行）。",
      "4. IDL 与 DL 员工入职满 6 个月后，可享受 5 天轮休假，且必须一次性休完。"
    ),
    keywords: ["年假", "轮休假", "职员年假单", "请假申请表", "JiaLian Li", "IDL", "DL"]
  },
  {
    rowNumber: 35,
    category: "人事/考勤",
    entryType: "流程",
    title: "其他假期",
    contacts: "JiaLian Li",
    body: joinLines(
      "婚假、产假、丧假、陪产假、育儿假、哺乳假、年假、无薪假等假期的申请流程和证明资料要求，请当面咨询 JiaLian Li。",
      "手册中还提示可参考以下文件：",
      "1. 雇员带薪年休假政策-更新202301",
      "2. 301区考勤政策",
      "3. 301雇员手册",
      "4. 301区婚假、丧假、产假、病假、医疗期工资支付政策更新-201609.pdf",
      "5. 育儿假及独生子女护理假的操作细则.pdf"
    ),
    keywords: ["其他假期", "婚假", "产假", "丧假", "陪产假", "育儿假", "无薪假", "JiaLian Li"]
  },
  {
    rowNumber: 36,
    category: "人事/考勤",
    entryType: "流程",
    title: "资料提交与审批规范（含时限、加签要求）",
    contacts: "JiaLian Li",
    body: joinLines(
      "1. 所有考勤相关资料（假期申请、加班表、转班、补卡等）均需 HR 加签。",
      "2. 每月 1 日下午 17:00 前，考勤员会发出各部门上月考勤汇总报表，由雇员本人确认签字后交考勤组存档。",
      "3. HR 考勤操作员每月 2 日 12:00 截止收取上月考勤资料。",
      "4. 员工需在当月 2 日上午前将考勤异常反馈给 JiaLian Li，否则可能影响工资核算。"
    ),
    keywords: ["考勤资料", "HR加签", "考勤汇总报表", "工资核算", "JiaLian Li"]
  },
  {
    rowNumber: 37,
    category: "人事/考勤",
    entryType: "流程",
    title: "江门厂区考勤管理",
    contacts: "JiaLian Li",
    body: joinLines(
      "江门厂区使用盖娅考勤管理系统，手册提示可参考《新入盖娅注意事项.pdf》了解手机端下载和注册方法。",
      "1. 江门同事的考勤问题原则上需自行在手机 App 上处理。",
      "2. 特殊情况（如转班、考勤异常、加班等）如超过 3 天未处理，需要本人发邮件给直属主管审批后，再通知 JiaLian Li 联系江门考勤员手动录入。",
      "3. 盖娅系统每月 1 日 19:00 关闭，关闭后无法录入上月考勤信息，可能导致考勤异常。",
      "4. JiaLian Li 会不定时检查盖娅系统考勤异常，并发邮件通知处理。"
    ),
    keywords: ["江门考勤", "盖娅", "Gaia", "转班", "加班", "JiaLian Li", "KM班车"]
  },
  {
    rowNumber: 38,
    category: "采购",
    entryType: "流程",
    title: "采购相关规定",
    relatedForm: "PR单",
    contacts: "JiaLian Li；YuLing Ye；ZhangBo Chu",
    body: joinLines(
      "因工作需求采购物品时，应先向相关采购员询问型号、规格、报价等信息，再按要求邮件发送给 JiaLian Li 或 YuLing Ye 开采购 PR 单。",
      "PR 模板关键字段包括：CAR no & Budget no、Description & BPA、Quantity、Unit of Measure、Unit Price、Currency、Buyer Name、Cost Center、User Name。",
      "补充规则：",
      "1. 单价超过 USD 5,000 时，必须提供 CAR no 和 Budget no。",
      "2. CAR# 可咨询 BU 财务 ZhangBo Chu（Ext 3405）。",
      "3. 如因项目紧急需自行采购物料，请参考“现金采购及报销流程”。",
      "4. 直接物料和非直接物料可通过链接查询。"
    ),
    url: "https://jehl.sharepoint.com/sites/JEInMotion/SitePages/SCS-Asia.aspx",
    keywords: ["采购", "PR单", "CAR", "Budget", "YuLing Ye", "ZhangBo Chu", "直接物料", "非直接物料"]
  },
  {
    rowNumber: 39,
    category: "报销/财务",
    entryType: "流程",
    title: "现金采购及报销流程",
    relatedForm: "EAR单",
    contacts: "YanNi Chen；Cally kf Chung；XiaoHong Zhou",
    body: joinLines(
      "1. 一般现金采购：1000 RMB 以内，经 Team Leader 批准后邮件通知相关采购员；若通过网上购买，需备注物品名称、数量和费用明细；经采购总监 Cally kf Chung 批准后方可自行购买。",
      "2. P-Card：5000 RMB 以内，一般用于 Non-budget 场景，直接找 PM 跟进。",
      "3. EAR 报销：一般适用于零件、夹具、模具等购买，大型设备不适用；经 XiaoHong Zhou 审批后可自行购买，再通过 EAR 单报销。",
      "4. 发票抬头需使用签订劳动合同上的公司名称，不确定时可找 YanNi Chen 确认。",
      "5. 两个常用开票信息如下：",
      "  - 德昌电机（深圳）有限公司，税号 91440300618901492T，地址深圳市宝安区新桥街道象山社区新发南路6号德昌工业园第三座2层（一照多址企业），电话 0755-29900656，开户行中国银行股份有限公司深圳市分行沙井支行，账号 744557933352，英文名 Johnson Electric (Shenzhen) Co., Ltd.",
      "  - 广东德昌电机有限公司，税号 914403007542779116，地址深圳市宝安区新桥街道象山社区新发南路6号德昌工业园第九座101（一照多址企业），电话 0755-29900656，开户行中国银行股份有限公司深圳市分行沙井支行，账号 760157933426，英文名 Johnson Electric (Guangdong) Co., Ltd.",
      "6. 发票要求：单位不能写“批”；若为“批”须提供税控机打印清单；货物名称不能写英文；发票内容不得留空，复核人与开票人不能为同一人；发票专用章需清晰且税号一致；发票不建议用胶水粘死，订书机装订即可。"
    ),
    keywords: ["现金采购", "报销", "EAR", "P-Card", "YanNi Chen", "Cally kf Chung", "XiaoHong Zhou", "开票资料"]
  },
  {
    rowNumber: 40,
    category: "报销/财务",
    entryType: "参考",
    title: "增值税发票开具规范",
    body: joinLines(
      "发票上所有信息内容必须完整，不能为空白。",
      "发票内容必须通过防伪税控系统开具，联次一次打印，不得手写。",
      "发票内容不得压线、超框，且必须加盖发票专用章。"
    ),
    keywords: ["增值税发票", "发票开具规范", "防伪税控", "发票专用章"]
  },
  {
    rowNumber: 41,
    category: "报销/财务",
    entryType: "流程",
    title: "报销相关规定（总则）",
    relatedForm: "TravelApprovalApp",
    contacts: "YanNi Chen；Jinn Hou；QianQian Yang",
    body: joinLines(
      "因公务出差、工作交际、礼物等发生的费用，如需报销，必须属于公务费用，并经成本中心主管审批。",
      "对接人：YanNi Chen。",
      "大陆同事报销会计负责人：Jinn Hou（Ext 3573）。",
      "香港同事报销会计负责人：QianQian Yang（Ext 2437）。",
      "出差报销需附成本中心主管批准的出差申请单（TravelApprovalApp - Power Apps）。"
    ),
    keywords: ["报销总则", "YanNi Chen", "Jinn Hou", "QianQian Yang", "TravelApprovalApp"]
  },
  {
    rowNumber: 42,
    category: "报销/财务",
    entryType: "流程",
    title: "住宿费报销",
    relatedForm: "EAR单",
    body: joinLines(
      "1. 客房费可凭旅馆盖章发票报销。",
      "2. 在中国大陆出差时，应优先入住公司协议酒店；如当地没有协议酒店，可自行预订。如未入住协议酒店，需在 EAR 系统备注原因。",
      "3. 电话费报销需在 EAR 里注明用途。",
      "4. 交通费用（火车票、汽车票、出租车费、地铁费等）均凭发票实报实销。"
    ),
    keywords: ["住宿费报销", "EAR", "协议酒店", "客房费", "电话费", "交通费"]
  },
  {
    rowNumber: 43,
    category: "报销/财务",
    entryType: "流程",
    title: "交通费报销",
    relatedForm: "EAR单",
    body: joinLines(
      "1. 应优先通过公司内部渠道预订机票和用车安排（公司订车限深圳附近地区）。",
      "2. 当公司车辆不能满足需求时，客运部会建议员工自行乘车，并凭发票报销。",
      "3. 交通费需在车票上注明始发点和终点站。",
      "4. 打车车型应选择滴滴快车或如祺出行；若乘坐专车/豪华商务车，必须在 EAR 中注明不乘坐快车的原因。"
    ),
    keywords: ["交通费报销", "EAR", "滴滴快车", "如祺出行", "用车安排"]
  },
  {
    rowNumber: 44,
    category: "报销/财务",
    entryType: "流程",
    title: "招待费和礼物费用报销",
    relatedForm: "EAR单",
    body: joinLines(
      "1. 招待费和礼物费用必须在 EAR 中注明礼物名称、客户姓名和职位、对方公司名称。",
      "2. 与本公司同事之间的餐费不能视为因公费用报销。",
      "3. 招待餐费需备注对方公司名称和参加人员姓名、职位；JE 内部人员还需提供工号和姓名。"
    ),
    keywords: ["招待费", "礼物费用", "EAR", "客户姓名", "公司名称", "餐费报销"]
  },
  {
    rowNumber: 45,
    category: "报销/财务",
    entryType: "系统链接",
    title: "单据提交时限与EAR路径",
    relatedForm: "EAR单",
    contacts: "YanNi Chen",
    body: joinLines(
      "1. 出差后请尽快完成报销流程，通常不超过 3 个月。",
      "2. 发票应按时间顺序和类别贴好，避免人员交接丢失。",
      "3. 建议先自行扫描留底，再将系统 EAR 订单页面打印出来，连同原始单据和发票一并交给 YanNi Chen 转财务审核。",
      "4. EAR 单填写路径见外部链接。"
    ),
    url: "https://www09.johnsonelectric.com/project/ears/ears.nsf",
    keywords: ["EAR", "报销时限", "YanNi Chen", "原始单据", "发票留底"]
  },
  {
    rowNumber: 46,
    category: "报销/财务",
    entryType: "系统链接",
    title: "报销描述说明和TRF样板",
    relatedForm: "TRF",
    body: joinLines(
      "手册页面给出了 TRF/报销描述的填写样板。",
      "样板提示包括：",
      "1. Transportation：写明起止地点以及必须乘坐出租车的原因。",
      "2. Accommodation fee：写明酒店名称。",
      "3. Express fee：写清寄送内容，以及寄件/收件地址信息。",
      "4. Meals fee / entertainment：说明是否与客户同行、客户姓名/公司名称以及招待原因；团队餐通常不建议走 EAR，因为流程较长且审批难度高。",
      "出差详细规定可参考外部链接。"
    ),
    url: "https://www09.johnsonelectric.com/project/ears/ears.nsf/ShowHelp.xsp",
    keywords: ["TRF", "报销描述", "Transportation", "Accommodation fee", "Express fee", "Meals fee", "EAR"]
  },
  {
    rowNumber: 47,
    category: "报销/财务",
    entryType: "流程",
    title: "Non-PO报销",
    contacts: "DanDan Song；Milly Ma",
    body: joinLines(
      "Non-PO 报销一般适用于书籍、报刊、杂志、测试标准文件、培训费用等特殊付款。",
      "流程：",
      "1. 发邮件给 DanDan Song，并提供费用描述、供应商名称、发票、相关合同或协议（需经法务审核、相关人员签字并加盖公章）等资料。",
      "2. DanDan Song 在 Non PO Payment - My Work 系统中填写并提交申请。",
      "3. 审批通过后，财务部门向供应商付款。",
      "4. 供应商必须已存在于公司供应商系统中，并有供应商代号和银行账号信息；如不在系统内，需要联系采购部 Milly Ma（Mobile 13632571345）新增供应商。"
    ),
    keywords: ["Non-PO", "DanDan Song", "Milly Ma", "Non PO Payment", "特殊付款", "供应商代号"]
  },
  {
    rowNumber: 48,
    category: "出差/后勤",
    entryType: "系统链接",
    title: "机票预定",
    relatedForm: "TravelApprovalApp",
    body: joinLines(
      "1. 先提交出差申请单 TravelApprovalApp - Power Apps，经成本中心主管审批后会生成 Travel Approval No。",
      "2. 拿到 Travel Approval No 后，可在携程商旅平台完成机票预订。",
      "3. 机票通常不需要员工个人垫付，平台由公司支付。"
    ),
    url: "https://ct.ctrip.com/singlesignon/openapi/saml/login/HUASHENGDIANJI",
    keywords: ["机票预定", "TravelApprovalApp", "Travel Approval No", "携程商旅", "Ctrip"]
  },
  {
    rowNumber: 49,
    category: "出差/后勤",
    entryType: "流程",
    title: "小车申请",
    relatedForm: "Transportation Service Application (Shenzhen & Jiangmen)",
    contacts: "YanNi Chen",
    body: joinLines(
      "1. 因公务需要用车时，应在用车前一天下午 2:00 前在系统“Shenzhen & Jiangmen 小车使用申请 / Transportation Service Application”中提交申请。",
      "2. 系统审批后，将申请单邮件转发给 YanNi Chen 安排司机。",
      "3. 司机信息一般会在用车前一天晚上 7 点左右发给乘车人。",
      "4. 如公司车辆不能满足需求，客运部会安排滴滴打车或通知自行乘车；滴滴费用由公司平台支付。",
      "5. 如自行搭车，必须选择有正规营业执照的车辆（公交车、客运站大巴、的士等），后续凭发票按 EAR 流程报销。"
    ),
    keywords: ["小车申请", "Transportation Service Application", "YanNi Chen", "滴滴", "EAR", "公务用车"]
  },
  {
    rowNumber: 50,
    category: "出差/后勤",
    entryType: "流程",
    title: "线路巴士",
    contacts: "LiHong Zhong；YanNi Chen",
    body: joinLines(
      "1. 公司为住在深圳市内及宝安区的雇员提供交通车，可按需向人力资源部申请。",
      "2. 具体路线可参考《大陆职员线路时刻表》，详情可咨询客运部 LiHong Zhong（Ext 3177）。",
      "3. 为方便公务往返 106、109、200 厂区，公司还安排了穿梭巴，可参考《穿梭巴时刻表》。",
      "4. 穿梭巴相关疑问可咨询行政组 YanNi Chen。"
    ),
    keywords: ["线路巴士", "交通车", "LiHong Zhong", "穿梭巴", "106", "109", "200", "YanNi Chen", "KM班车"]
  },
  {
    rowNumber: 51,
    category: "出差/后勤",
    entryType: "流程",
    title: "穿梭巴士预定",
    relatedForm: "Shuttle/Checkpoint Bus Reservation System；小车申请表",
    body: joinLines(
      "1. 用车前一天 17:00 前，在 Shuttle/Checkpoint Bus Reservation System 中提交 KM、口岸巴士用车申请。",
      "2. 临时用车需手写《小车申请表》交给车队。",
      "3. 特别提醒：周二和周四早上没有从 KM 至 P200 的穿梭巴；周二和周四下午没有从 P200 至 KM 的穿梭巴。"
    ),
    keywords: ["穿梭巴", "Shuttle/Checkpoint Bus Reservation System", "KM", "P200", "口岸巴士", "小车申请表"]
  },
  {
    rowNumber: 52,
    category: "出差/后勤",
    entryType: "参考",
    title: "职员线路车时刻表",
    body: joinLines(
      "雇员上下班交通车时间表（2024 年 03 月更新）主要线路如下：",
      "1. A1 线：7:20 兴东地铁口发车，经海天居、友谊书城、前进税务局、诺铂广场、沃尔玛、富盈门、天骄世家、丽景城到 P200；17:45 从 P200 返程。",
      "2. A2 线：7:33 北方公司发车，经海都小学、奉华明珠、港隆城、翠景居、海城派出所、桃景居到 P200；17:45 从 P200 返程。",
      "3. B2 线：7:25 尚都花园发车，经海岸花园、海月华庭、中港、财富港、海福一、幸福海岸、尚都花园到 P200；17:45 从 P200 返程。",
      "4. C2 线：7:15 南油公交站发车，经同心海雅、西城上筑、波尔诺、滨海春城等到 P200；17:45 从 P200 返程。",
      "5. D2 线：7:00 梅华小学发车，经翻身、甲岸、灵芝公园、英伦名苑、松坪山公交站、宝安中学等到 P200；17:45 从 P200 返程。",
      "6. E2 线：7:10 五和地铁口发车，经万众城、锦绣江南、上芬地铁、电信大厦、美丽家园、工商所、清湖路口到 P200；17:45 从 P200 返程。",
      "7. F 线：7:00 宝安北路发车，经太白路口、荣超花园、龙珠花园、木棉湾、东方半岛、布吉警署等到 P200；17:45 从 P200 返程。",
      "8. 沙井线：7:40 金沙市场发车，经井金沙市场公交站、南埔苑公交站、万丰天桥公交站、市民广场创新天虹/赛丰路、上星社区公交站、至尊大门口到 P200；17:45 从 P200 返程。",
      "备注：以首发站时间为准，中间站点仅为参考时间，考虑道路交管事实，请中间站点同事提前候车，原则上车到即走，不候车。"
    ),
    keywords: ["线路车时刻表", "A1线", "A2线", "B2线", "C2线", "D2线", "E2", "F线", "沙井线", "P200"]
  },
  {
    rowNumber: 53,
    category: "出差/后勤",
    entryType: "参考",
    title: "来往厂穿梭巴时刻表",
    body: joinLines(
      "301 穿梭巴时间表适用于周一至周五，节假日除外；来往分厂穿梭巴附时间表，公务往返分厂的雇员请选择穿梭巴。",
      "德昌大厦 -> P200 发车时间：6:35、6:40、7:40、8:05、8:40、9:00、10:00、11:00、11:35、12:05、13:05、14:00、15:00、16:20、17:05、18:10、18:30。",
      "P200 -> 德昌大厦 回程时间：7:00、7:40、8:20、8:30、9:00、10:00（备注：经过 109）、11:00、11:45、12:45、13:05、14:00、15:00、16:40、17:45、18:05、18:45、19:05、19:15。",
      "P200 <-> P109：",
      "1. 8:30 线船车：200 -> 109；回程 8:50。",
      "2. 8:50 口岸车：200 -> 109。",
      "3. 10:00：200 -> 109；回程 10:30。",
      "4. 16:00 口岸车：200 -> 109；回程 16:40（口岸车乘客专用车）。",
      "5. 17:05：109 -> 200（线路车乘客专用车）。",
      "备注：蓝色部分为 P200 班次，橙色部分为 P109 班次；本时间表自 2024-05-10 起生效。"
    ),
    keywords: ["穿梭巴时刻表", "301 Shuttle Bus", "德昌大厦", "P200", "P109", "口岸车", "线船车"]
  },
  {
    rowNumber: 54,
    category: "出差/后勤",
    entryType: "流程",
    title: "货车申请",
    contacts: "YanNi Chen；江保杰",
    body: joinLines(
      "1. 提前一天发邮件给 YanNi Chen 申请，需提供物品名称、数量、尺寸、使用日期、时间、装车和卸货地点、物品接收人及手机号码等信息。",
      "2. 也可在 Transport Application 系统中填写信息提交申请，待成本中心主管审批后，由公司货运部安排合适大小的货车。",
      "3. 负责人：江保杰，Ext 3276，Mobile 13715183416。",
      "4. 公司每天固定安排 KM-P200 往返货车；若需使用，请联系 YanNi Chen。"
    ),
    keywords: ["货车申请", "Transport Application", "YanNi Chen", "江保杰", "KM", "P200"]
  },
  {
    rowNumber: 55,
    category: "出差/后勤",
    entryType: "流程",
    title: "国内快递与国外快递",
    relatedForm: "DHL Shipment Form；Non-Oracle Shipment Application；PFA；物品出门纸",
    body: joinLines(
      "国内快递：到行政组拿顺丰单扫码填写相关信息 -> 填写出门纸 -> 扫 JE 码登记。",
      "国外快递：填写《DHL Shipment Form》 -> 填写物品出门纸 -> 顺丰单扫码填写寄件信息 -> 扫 JE 码登记 -> 在 Non-Oracle Shipment Application 系统填写信息；超过 30 公斤的货物还需填写 PFA。"
    ),
    keywords: ["国内快递", "国外快递", "顺丰", "DHL Shipment Form", "Non-Oracle Shipment Application", "PFA", "物品出门纸"]
  },
  {
    rowNumber: 56,
    category: "出差/后勤",
    entryType: "参考",
    title: "DHL快递特别提醒",
    body: joinLines(
      "1. DHL 公司账号 Payer Account No.：630873225。寄出时将该账号填在 DHL 快递单上；寄进时将该账号告知寄件人，由国外分公司快递中心寄出。",
      "2. 需填写的内容包括：寄件人地址、收件人地址、寄送物品中英文名称、物品型号、打包尺寸、货物总价/单价/币种、净重/毛重、寄送理由。",
      "3. 不能寄送整个马达、粉末和油性物品。",
      "4. 对于有磁性的产品，应逐个用泡泡纸包装，再整体放入泡泡塑料箱，外层再用纸箱包装。"
    ),
    keywords: ["DHL", "Payer Account No", "630873225", "磁性产品", "快递提醒"]
  },
  {
    rowNumber: 57,
    category: "出差/后勤",
    entryType: "流程",
    title: "沙井-江门往返物品寄送",
    relatedForm: "Electronic Cargo Transfer Process",
    body: joinLines(
      "1. 沙井-江门往返物品需经货车寄出，先在 Electronic Cargo Transfer Process 系统中提交寄件人和收件人姓名、电话等信息。",
      "2. 成本中心主管审批后，单面打印 2 份，并与打包好的物品一起，在当天下午 14:00 前放到行政组寄出。",
      "3. 私人物品不可以在公司邮递。"
    ),
    keywords: ["江门往返寄送", "货车寄送", "Electronic Cargo Transfer Process", "寄件人", "收件人"]
  },
  {
    rowNumber: 58,
    category: "出差/后勤",
    entryType: "参考",
    title: "物品出门纸填写模板",
    body: joinLines(
      "手册中提供了物品出门纸填写模板。",
      "备注：有签字权限的一般为部门主管、总监及项目经理。",
      "特别提醒：私人笔记本电脑/相机不可以带进厂区。"
    ),
    keywords: ["物品出门纸", "填写模板", "签字权限", "私人笔记本", "相机"]
  },
  {
    rowNumber: 59,
    category: "出差/后勤",
    entryType: "流程",
    title: "深圳&江门宿舍预定",
    relatedForm: "Accommodation Application (Shenzhen & Jiangmen)",
    body: joinLines(
      "1. 在“深圳&江门住宿申请 / Accommodation Application (Shenzhen & Jiangmen)”系统中填写并提交申请。",
      "2. 审批完成后由宿舍服务处安排。",
      "3. 入住当天到宿舍服务处领取钥匙。",
      "宿舍服务处联系方式：",
      "  - SZ Dormitory 深圳宿舍：(0755) 29900701 / 816-7235；p200.dormitory.service@johnsonelectric.com",
      "  - SZ Hostel 深圳旅馆：(0755) 29853388 / 816-3522；hostel.106@johnsonelectric.com",
      "  - JM Dormitory 江门宿舍：(0750) 3202008 / 817-2008；jm.dormitory@johnsonelectric.com",
      "  - JM Hostel 江门旅馆：(0750) 3202888 / 817-2888；jm.hostel@johnsonelectric.com"
    ),
    keywords: ["宿舍预定", "Accommodation Application", "深圳宿舍", "江门宿舍", "hostel", "dormitory"]
  },
  {
    rowNumber: 60,
    category: "出差/后勤",
    entryType: "参考",
    title: "江门厂区餐补",
    body: joinLines(
      "1. 新入职或第一次去江门出差时，需要先在 C1-1F 食堂文书处激活厂牌，用于吃饭刷卡抵扣。",
      "2. 每天可刷两次补贴：中午 11:30-13:30，下午 17:30-20:30，每次抵扣 7.5 元。",
      "3. 若未在食堂刷卡，也可在 D1-1F 小卖部消费抵扣；小卖部刷卡时间为 11:30-13:30、16:30-20:30。",
      "4. 职员请假期间及周末休息没有餐补。"
    ),
    keywords: ["江门餐补", "厂牌激活", "食堂", "小卖部", "7.5元"]
  },
  {
    rowNumber: 61,
    category: "行政支持",
    entryType: "流程",
    title: "会议室的预定与管理",
    contacts: "YanNi Chen；Wei Guo",
    body: joinLines(
      "1. 在会议室系统中预定，系统名称为 Corp Eng Meeting Room Reservation - Power Apps。",
      "2. 没有权限或不会预定时，可联系管理员：P200-3B-2F 为 YanNi Chen，KM-C1-4F 为 Wei Guo。",
      "3. 使用后应保持会议室干净整齐，爱护室内设备和公用设施。",
      "4. 离开前需将桌椅归位、白板擦净，并关闭电源、投影仪和空调。",
      "5. 如未遵守要求，管理员有权拒绝该预订者的下次预订请求。"
    ),
    keywords: ["会议室预定", "Power Apps", "YanNi Chen", "Wei Guo", "投影仪", "白板"]
  },
  {
    rowNumber: 62,
    category: "资产/EHS",
    entryType: "流程",
    title: "资产转移、报废",
    contacts: "JiaLian Li；DanDan Song",
    body: joinLines(
      "资产转移：使用人或成本中心转移时，经成本中心批准后发邮件给 JiaLian Li 开资产转移单。",
      "资产转卖：如 JEGD 转卖到 HSGD，需按内部公司资产转移流程完成账务买卖。",
      "资产报废：提供资产号，发邮件给成本中心主管并抄送 JiaLian Li；获得主管邮件批准后由 JiaLian Li 开资产报废单。报废单批完后，将待报废资产送到报废仓。",
      "非资产（如办公桌椅）、马达、零件、电路板等报废需分类整理并拍照，发给 DanDan Song 安排处理。"
    ),
    keywords: ["资产转移", "资产报废", "JiaLian Li", "DanDan Song", "报废仓", "资产号"]
  },
  {
    rowNumber: 63,
    category: "资产/EHS",
    entryType: "流程",
    title: "仪器工具转移",
    relatedForm: "工具仪器转移单；工具仪器改动表",
    contacts: "JiaLian Li",
    body: joinLines(
      "1. 仪器工具转移时，将双方成本中心主管签字的《工具仪器转移单》交给 JiaLian Li 办理。",
      "2. 特殊资产/仪器事项（如遗失）由负责人整理并说明原因，再由 JiaLian Li 与负责人共同处理。",
      "3. 使用人离职时，需先找到资产/仪器位置，再由申请人填写《工具仪器改动表》，经成本中心主管签字后交 JiaLian Li。",
      "4. 历史遗留资产/仪器也需由负责人整理原因，JiaLian Li 确认后再安排处理。"
    ),
    keywords: ["仪器工具转移", "工具仪器转移单", "工具仪器改动表", "JiaLian Li", "仪器遗失"]
  },
  {
    rowNumber: 64,
    category: "资产/EHS",
    entryType: "流程",
    title: "仪器检定",
    contacts: "JiaLian Li",
    body: joinLines(
      "1. 检定分为内部检定、现场检定和外送检定。",
      "2. 部门负责人 JiaLian Li 一般在每月 1 号左右发出需要检定的物品检定单。",
      "3. 内部检定结束后应及时取回仪器。",
      "4. 所有外检仪器必须在检定有效期到期前送出，或备注为现场检定，严禁过期未送检。"
    ),
    keywords: ["仪器检定", "内部检定", "现场检定", "外送检定", "JiaLian Li"]
  },
  {
    rowNumber: 65,
    category: "安防/访客",
    entryType: "流程",
    title: "供应商进厂申请",
    relatedForm: "外来人员入厂工作申请表",
    body: joinLines(
      "流程：填写《外来人员入厂工作申请表》 -> 部门主管审核签字 -> 供应商在北门登记 -> 接待人签字 -> 将申请表交北门保安 -> 保安分发入厂证并放行。"
    ),
    keywords: ["供应商进厂", "外来人员入厂工作申请表", "北门", "入厂证"]
  },
  {
    rowNumber: 66,
    category: "安防/访客",
    entryType: "流程",
    title: "客户来访申请",
    relatedForm: "FVP系统（Facility Visit Program）",
    contacts: "DanDan Song",
    body: joinLines(
      "1. 需至少提前 3 天在 FVP 系统 Facility Visit Program 中提交申请，填写客户公司名称、客户姓名、职位、来访理由等。",
      "2. 客户来访如需会议室，一般选择 P200 9 座 1 楼前台处。",
      "3. FVP 单审批完成后，凭单号到 P200 南门邮递中心领取访客 VIP 证。",
      "4. 如未提前 3 天申请，系统会提示不能提交；可先保存信息后截图发给 DanDan Song，请管理员后台处理。"
    ),
    keywords: ["客户来访", "FVP", "Facility Visit Program", "VIP证", "DanDan Song", "南门邮递中心"]
  },
  {
    rowNumber: 67,
    category: "安防/访客",
    entryType: "流程",
    title: "沙井厂区笔记本电脑/相机出门",
    relatedForm: "笔记本电脑/相机入闸纸；笔记本电脑/相机出闸纸",
    contacts: "King Tang；武小辉",
    body: joinLines(
      "1. 先由乙方提供品牌、型号、电脑编号等 3 个不同号码。",
      "2. 填写《笔记本电脑/相机入闸纸》和《笔记本电脑/相机出闸纸》，由被访部门主管审批。",
      "3. 再到 10 座 3 楼电脑部由 King Tang 审批。",
      "4. 再到 9 座天台安防部由武小辉审批。",
      "5. 在保安处登记后放行；供应商出厂时需携带《笔记本电脑/相机出闸纸》。"
    ),
    keywords: ["笔记本出门", "相机出门", "入闸纸", "出闸纸", "King Tang", "武小辉", "沙井厂区"]
  },
  {
    rowNumber: 68,
    category: "安防/访客",
    entryType: "流程",
    title: "江门厂区笔记本电脑/相机出门",
    relatedForm: "笔记本电脑/相机入闸纸；笔记本电脑/相机出闸纸",
    contacts: "刘辉；王平",
    body: joinLines(
      "1. 先由乙方提供品牌、型号、电脑编号等 3 个不同号码。",
      "2. 填写《笔记本电脑/相机入闸纸》和《笔记本电脑/相机出闸纸》，由被访部门主管审批。",
      "3. 到江门接待处由刘辉执行杀毒。",
      "4. 再由安防部王平审批。",
      "5. 在保安处登记后放行；供应商出厂时需携带《笔记本电脑/相机出闸纸》。"
    ),
    keywords: ["笔记本出门", "相机出门", "入闸纸", "出闸纸", "刘辉", "王平", "江门厂区"]
  },
  {
    rowNumber: 69,
    category: "安防/访客",
    entryType: "流程",
    title: "江门访客系统相关流程",
    body: joinLines(
      "1. 供应商需提前关注德昌电机服务号，点击“访客申请”，完成注册并填写相关资料。",
      "2. 待成本中心主管审批完成后方可入厂。",
      "3. 供应商只能从 3 号门进厂。",
      "4. 如携带电脑，还需额外提交笔记本电脑进出申请，并在接待处杀毒后方可进入。"
    ),
    keywords: ["江门访客", "访客申请", "服务号", "3号门", "笔记本电脑进出申请"]
  },
  {
    rowNumber: 70,
    category: "资产/EHS",
    entryType: "参考",
    title: "6S定义",
    contacts: "Jie Li",
    body: joinLines(
      "6S 定义页为图文海报，以下按可辨识内容转录；少量细字仍建议结合原图人工复核。",
      "1. 整理（Structurise / Set Aside Unused Space）：",
      "   - 定义：将物品分为必要和不必要的，不必要的物品要尽快处理掉。",
      "   - 目的：腾出空间、空间活用、防止误用误送，并塑造清爽的工作场所。",
      "   - 推行要领：仓储检查工作场所（含看得见和看不见处）；制定“要/不要”判别基准；清除不必要品；调查使用频率并决定日常用量和放置位置；制定废弃物处理办法并每日自检。",
      "2. 整顿（Systematise / Enhance Work Efficiency）：",
      "   - 定义：必要物品分门别类放置、排列整齐、规定数量，并做有效标识。",
      "   - 目的：让工作场所一目了然、整齐有序，减少找寻物品时间，并识别过多积压物品。",
      "   - 推行要领：完成整理工作；明确必要物品放置场所；摆放整齐有条不紊；地板划线定位；场所和物品标识清楚；制订废弃物处理办法。",
      "3. 清洁（Sanitise / Remove Debris & Maintain Fresh And Clean Environment）：",
      "   - 定义：将工作场所清扫干净，保持工作场所干净、亮丽。",
      "   - 目的：清除脏污，保持工作场所干净明亮，从而稳定品质、减少工业危害。",
      "   - 推行要领：建立清洁责任区域；实施例行扫除和清理脏污；调查污染源并杜绝或隔离；建立清洁基准作为规范；定期开展全面大清洁。",
      "4. 规范（Standardise / Maintain Achievement）：",
      "   - 定义：维持整理、整顿、清洁所取得的成果，并使之标准化、制度化。",
      "   - 目的：维持前面 3S 的成果。",
      "   - 推行要领：落实前面 3 个“S”，制订目视管理基准、6S 实施办法、考评和稽核办法、奖惩制度，加强执行，并由经理经常带头巡查，带动全员重视 6S 活动。",
      "5. 安全（Safety / Abide by Rules / Crush in the egg）：",
      "   - 定义：采取系统措施保证人员、场地、物品的安全。",
      "   - 目的：系统建立防烫伤、防污、防火、防水、防盗、防损等安全措施，营造无事故隐患环境，避免安全事故苗头，减少工业灾害。",
      "   - 推行要领：通过实施整理、整顿、清洁、规范，创造安全作业环境，做到预防为主、防微杜渐。",
      "6. 自律（Self-discipline / Maintain High Discipline）：",
      "   - 定义：通过宣传、培训、激励等方法，把外在管理要求转化为员工自身习惯意识，使各项活动成为发自内心的自觉行动。",
      "   - 目的：建立习惯与意识，从根本上提升人员修养品质，使员工对任何工作都讲究认真。",
      "   - 推行要领：以人为出发点，通过整理、整顿、清洁、规范、安全等合理化改善活动培养共同管理语言，使全员养成良好习惯，遵守规范和标准，进一步提升全面管理水平。"
    ),
    keywords: ["6S", "整理", "整顿", "清洁", "规范", "安全", "自律", "Jie Li"]
  },
  {
    rowNumber: 71,
    category: "资产/EHS",
    entryType: "流程",
    title: "EHS相关要求",
    contacts: "Jie Li",
    body: joinLines(
      "1. 每月 28 号左右召开安全月度会议并开展安全检查，发现问题要及时整改，各部门平时也需自检、排除安全隐患。",
      "2. 有噪音、粉尘、异味的施工应安排在非正常上班时间（周末、节假日等）；各安全主任需监督辖区内施工，发现异常或事故应立即告知行政组，并协同调查处理。",
      "3. 施工结束后，应要求施工人员清理现场垃圾，并通知行政组安排保洁。",
      "4. 危险废弃物（废油、含油废水、含油抹布、废油墨、废电池等）严禁当普通垃圾丢弃，必须交由部门安全主任送至集团工程部危废仓处理。",
      "5. 特殊环境需佩戴防护用品：异味/灰尘环境戴专用口罩，噪音环境戴耳塞，经常去车间或搬运重物时穿工鞋。",
      "6. 搬运大件仪器/设备时，可请专业搬运公司，避免工伤。",
      "7. 如购买新机器，应在到货前 3 个工作日通知 Jie Li 开展验机工作，以加强装调机阶段安全风险控制。",
      "8. EHS 资料参考路径：N:\\CHN\\ENG\\ENG_PUB\\Corp Eng\\Corp.Eng.EHS"
    ),
    keywords: ["EHS", "危废", "工鞋", "口罩", "耳塞", "验机", "Jie Li", "安全月度会议"]
  },
  {
    rowNumber: 72,
    category: "资产/EHS",
    entryType: "参考",
    title: "SafetyCulture安全文化平台",
    contacts: "Jie Li",
    body: joinLines(
      "SafetyCulture 平台用于帮助员工更快速地传达观察到的 EHS 问题和担忧，并让管理团队跟进闭环。",
      "平台定位：为本地领导团队提供管理和解决 EHS 问题的系统，也为 EHS 人员提供多种集成工具，以提升效率和结果交付。",
      "使用方式：手册页附有二维码，发现安全隐患时可扫码提报；EHS 部门会对提交隐患的人员发放小奖品以示激励。",
      "如有疑问，可咨询 Jie Li。"
    ),
    keywords: ["SafetyCulture", "安全文化平台", "EHS", "隐患提报", "Jie Li", "二维码"]
  }
];

function buildImageAttachments() {
  return rows.flatMap<ImageAttachment>((row) => {
    const pages = imagePagesByRow.get(row.rowNumber) ?? [];

    return pages.map((pageNumber) => {
      const fileName = `row-${pad2(row.rowNumber)}-page-${pad2(pageNumber)}.png`;
      return {
        rowNumber: row.rowNumber,
        pageNumber,
        fileName,
        relativePath: `attachments/${fileName}`,
        attachmentType: "png",
        description: `来源 PDF 页码：第 ${pageNumber} 页；关联条目：${row.title}`
      };
    });
  });
}

function materializeRows(imageAttachments: ImageAttachment[]) {
  const imageLinksByRow = new Map<number, string[]>();

  for (const attachment of imageAttachments) {
    const current = imageLinksByRow.get(attachment.rowNumber) ?? [];
    current.push(attachment.relativePath);
    imageLinksByRow.set(attachment.rowNumber, current);
  }

  return rows.map((row) => ({
    ...row,
    imageLinks: imageLinksByRow.get(row.rowNumber) ?? []
  }));
}

function createWorkbookBuffer(materializedRows: CanonicalRow[], imageAttachments: ImageAttachment[]) {
  const workbook = XLSX.utils.book_new();

  const versionSheet = XLSX.utils.aoa_to_sheet([
    ["版本", "日期", "说明", "作者"],
    ["2026.04.01", "2026.04.01", "由 Corp. Eng New Staff Guide Book PDF 转制", "集团工程部行政组"]
  ]);

  const flowSheet = XLSX.utils.aoa_to_sheet([
    ["编号", "一级分类", "条目类型", "标题", "相关单据", "联系人/责任人", "正文", "外部链接", "关键词/别名", "图片链接"],
    ...materializedRows.map((row) => [
      row.rowNumber,
      row.category,
      row.entryType,
      row.title,
      row.relatedForm ?? "",
      row.contacts ?? "",
      row.body,
      row.url ?? "",
      row.keywords.join("；"),
      row.imageLinks?.join("；") ?? ""
    ])
  ]);

  const attachmentSheet = XLSX.utils.aoa_to_sheet([
    ["条目编号", "附件文件名", "相对路径", "附件类型", "附件说明"],
    ...imageAttachments.map((attachment) => [
      attachment.rowNumber,
      attachment.fileName,
      attachment.relativePath,
      attachment.attachmentType,
      attachment.description
    ])
  ]);

  XLSX.utils.book_append_sheet(workbook, versionSheet, "版本说明");
  XLSX.utils.book_append_sheet(workbook, flowSheet, "流程汇总");
  XLSX.utils.book_append_sheet(workbook, attachmentSheet, "附件清单");

  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

async function renderImageAssets(imageAttachments: ImageAttachment[]) {
  if (!imageAttachments.length) {
    return;
  }

  await fs.access(sourcePdfPath);
  const uniquePages = [...new Set(imageAttachments.map((attachment) => attachment.pageNumber))].sort((left, right) => left - right);
  const tempDir = await fs.mkdtemp(path.join(tmpdir(), "corp-eng-guide-pages-"));
  const pagePrefix = path.join(tempDir, "page");

  try {
    await execFileAsync("pdftoppm", [
      "-r",
      "150",
      "-f",
      String(uniquePages[0]),
      "-l",
      String(uniquePages[uniquePages.length - 1]),
      "-png",
      sourcePdfPath,
      pagePrefix
    ]);

    for (const attachment of imageAttachments) {
      const renderedPagePath = path.join(tempDir, `page-${pad2(attachment.pageNumber)}.png`);
      await fs.copyFile(renderedPagePath, path.join(outputDir, attachment.relativePath));
    }
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

async function writeZip(workbookBuffer: Buffer, imageAttachments: ImageAttachment[]) {
  const zip = new JSZip();
  zip.file("knowledge.xlsx", workbookBuffer);

  if (!imageAttachments.length) {
    zip.folder("attachments");
  } else {
    for (const attachment of imageAttachments) {
      const buffer = await fs.readFile(path.join(outputDir, attachment.relativePath));
      zip.file(attachment.relativePath, buffer);
    }
  }

  const zipBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(zipPath, zipBuffer);
}

async function ensureOutput(workbookBuffer: Buffer, imageAttachments: ImageAttachment[]) {
  await fs.rm(outputDir, { recursive: true, force: true });
  await fs.mkdir(path.join(outputDir, "attachments"), { recursive: true });
  await renderImageAssets(imageAttachments);
  await fs.writeFile(path.join(outputDir, "knowledge.xlsx"), workbookBuffer);
  await writeZip(workbookBuffer, imageAttachments);
}

async function smokeTest(workbookBuffer: Buffer, materializedRows: CanonicalRow[], imageAttachments: ImageAttachment[]) {
  const workbook = XLSX.read(workbookBuffer, { type: "buffer" });
  const flowSheetRows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets["流程汇总"], {
    header: 1,
    blankrows: false
  });
  const attachmentSheetRows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets["附件清单"], {
    header: 1,
    blankrows: false
  });

  if (flowSheetRows[0]?.[9] !== "图片链接") {
    throw new Error("流程汇总 sheet 缺少 J 列“图片链接”。");
  }
  if (attachmentSheetRows.length - 1 !== imageAttachments.length) {
    throw new Error(`附件清单记录数异常：期望 ${imageAttachments.length}，实际 ${attachmentSheetRows.length - 1}`);
  }

  for (const attachment of imageAttachments) {
    await fs.access(path.join(outputDir, attachment.relativePath));
  }

  const requiredImageRows = [2, 3, 52, 53, 70];
  const missingRequiredRows = requiredImageRows.filter((rowNumber) => !(materializedRows.find((row) => row.rowNumber === rowNumber)?.imageLinks?.length));
  if (missingRequiredRows.length) {
    throw new Error(`以下条目缺少图片链接：${missingRequiredRows.join(", ")}`);
  }

  const multiImageRow = materializedRows.find((row) => row.rowNumber === 38);
  if (!multiImageRow?.imageLinks?.join("；").includes("；")) {
    throw new Error("多图条目的图片链接未按全角分号拼接。");
  }

  const parsed = await parseWorkflowWorkbook(workbookBuffer, "smoke-test", "knowledge.xlsx", {
    canonicalAttachmentsDir: path.join(outputDir, "attachments")
  });
  if (parsed.warnings.length > 0) {
    throw new Error(`解析出现 warning：${parsed.warnings.join(" | ")}`);
  }

  const keywordsToCheck = ["ITPM", "EAR", "FVP", "TravelApprovalApp", "KM班车"];
  const keywordCoverage = keywordsToCheck.map((keyword) => ({
    keyword,
    matched: rows.some((row) => `${row.title}\n${row.body}\n${row.keywords.join("\n")}`.includes(keyword))
  }));

  const missing = keywordCoverage.filter((item) => !item.matched);
  if (missing.length > 0) {
    throw new Error(`关键词覆盖失败：${missing.map((item) => item.keyword).join(", ")}`);
  }

  return {
    sourceCount: parsed.sources.length,
    warningCount: parsed.warnings.length,
    keywordCoverage,
    imageAttachmentCount: imageAttachments.length
  };
}

async function main() {
  const imageAttachments = buildImageAttachments();
  const materializedRows = materializeRows(imageAttachments);
  const workbookBuffer = createWorkbookBuffer(materializedRows, imageAttachments);
  await ensureOutput(workbookBuffer, imageAttachments);
  const result = await smokeTest(workbookBuffer, materializedRows, imageAttachments);

  console.log(`Generated ${rows.length} rows.`);
  console.log(`Image attachments: ${result.imageAttachmentCount}`);
  console.log(`Workbook: ${path.join(outputDir, "knowledge.xlsx")}`);
  console.log(`Zip: ${zipPath}`);
  console.log(`Parsed sources: ${result.sourceCount}`);
  console.log(`Warnings: ${result.warningCount}`);
  for (const item of result.keywordCoverage) {
    console.log(`Keyword ${item.keyword}: ${item.matched ? "ok" : "missing"}`);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
