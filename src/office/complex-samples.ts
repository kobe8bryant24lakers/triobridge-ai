import type { HostKind, ReportContent, SlideContent, TableContent } from '../core/types';
import { reportToText, tableToText } from '../core/structured-output';

export type DemoScenario = 'basic' | 'complex';

// Deterministic, synthetic fixtures shared by the browser demo and regression tests.
export const COMPLEX_REPORT: ReportContent = {
  title: '跨团队项目验收报告',
  sections: [
    {
      heading: '验收结论与范围',
      paragraphs: [
        '本报告为合成演示样稿。项目分为需求梳理、数据迁移、使用培训和交付验收四个阶段。当前已完成 48 项检查，其中 42 项通过、4 项需要复核、2 项等待外部资料。未关闭事项不计入已通过项目。',
        '本轮验收覆盖三个区域、四个交付小组和两类使用场景。附件、流程图与原始数据应在正式交付前逐项核对，不能用摘要代替附件检查。',
      ],
      bullets: [
        '通过率按 42 ÷ 48 计算，为 87.5%。',
        '上线决定需以全部必需事项关闭后的复核结果为准。',
      ],
    },
    {
      heading: '指标口径与结果',
      paragraphs: [
        '所有数量均为演示数据。响应时间使用小时，完成率使用通过检查数与总检查数的比值；平均值不能代替每个小组的达标情况。',
      ],
      bullets: [],
      table: {
        headers: ['指标', '期初', '本期', '复核说明'],
        rows: [
          ['平均响应时间', '24 小时', '8 小时', '按相同时间窗口比较'],
          ['验收检查', '48 项', '42 项通过', '4 项复核，2 项待资料'],
          ['通过率', '待补充', '87.5%', '分母包含未关闭事项'],
        ],
      },
    },
    {
      heading: '数据迁移与一致性',
      paragraphs: Array.from(
        { length: 4 },
        (_, i) =>
          `第 ${i + 1} 批数据需要核对记录标识、字段类型、关联关系与时间范围。保留空白和零值之间的区别，百分比应按原始数值计算后再格式化显示。带换行、制表符、中文、英文或特殊符号的单元格必须完整保留；如出现重复记录，应先确认业务含义再决定是否合并。此段为合成的长段落，用于检查内容阅读与编辑。`,
      ),
      bullets: ['抽样不能替代关键字段的完整校验。', '缺失资料单独登记，负责人填写后再次复核。'],
    },
    {
      heading: '风险与待办',
      paragraphs: [
        '以下事项尚未关闭，状态变化需要负责人确认。日期与负责人保留为占位信息，避免演示数据被误认为实际承诺。',
      ],
      bullets: [],
      table: {
        headers: ['事项', '负责人', '状态', '完成条件'],
        rows: [
          ['4 项复核', '待补充', '待复核', '复核通过并保存证据'],
          ['2 项外部资料', '待补充', '待提供', '资料到齐且内容一致'],
          ['使用培训', '待补充', '待确认', '记录参加情况和遗留问题'],
        ],
      },
    },
    {
      heading: '跨团队交接安排',
      paragraphs: [
        '每个交付小组整理本组输入、交付物、依赖事项和验收记录。接收方核对文件版本与内容后确认接收；发现差异时保留原始记录，补齐说明后再更新。',
        '对跨区域协作，明确时间口径与责任边界。工作表中的公式引用需包含准确的区域和期间，不能将多期累计值重复相加。',
      ],
      bullets: ['交接材料包含交付清单与未关闭事项。', '接收记录与补充资料保持可追溯。'],
    },
    {
      heading: '后续行动与复核',
      paragraphs: [
        '先关闭待复核事项，再更新验收结论。正式版应补齐负责人、时间和证据链接；本演示不包含真实个人信息或客户数据。',
      ],
      bullets: [
        '逐项补齐资料和验收证据。',
        '复算关键指标并检查格式。',
        '保留原始文件，在确认后提交新版本。',
      ],
    },
  ],
};

export const COMPLEX_TABLE: TableContent = {
  name: '区域运营演示',
  headers: ['月份', '区域', '项目', '收入', '成本', '差额', '差额率', '备注'],
  formats: ['text', 'text', 'text', 'number', 'number', 'number', 'percent', 'text'],
  rows: Array.from({ length: 48 }, (_, i) => {
    const row = i + 2;
    return [
      `${Math.floor(i / 4) + 1}月`,
      ['华东', '华南', '华北', '西南'][i % 4],
      `项目 ${String(i + 1).padStart(3, '0')}`,
      1200 + i * 37.25,
      800 + i * 19.125,
      { formula: `=D${row}-E${row}` },
      { formula: `=IFERROR(F${row}/D${row},0)` },
      i === 0 ? '合成数据\n保留两行备注' : i === 47 ? '末行核对' : null,
    ];
  }),
};

export const COMPLEX_SLIDES: SlideContent[] = [
  {
    title: '项目验收与下一阶段安排',
    bullets: [
      '合成演示数据：48 项检查，42 项通过',
      '4 项需要复核，2 项等待外部资料',
      '重点说明验收口径、未关闭事项与责任交接',
    ],
  },
  ...[
    '验收范围',
    '需求梳理',
    '数据迁移',
    '数值与公式校验',
    '区域协作',
    '培训与使用反馈',
    '交付材料',
    '风险与依赖',
    '后续行动',
    '复核机制',
  ].map((title) => ({
    title,
    bullets: [
      '以原始资料为依据，保留事实、数字和单位',
      '区分已完成事项与等待补充的信息',
      '明确负责人、完成条件与复核时间',
      '缺失信息用“待补充”标示，不代替实际确认',
    ],
  })),
  {
    title: '验收确认与行动清单',
    bullets: [
      '关闭全部必需事项后更新结论',
      '核对文档、表格及演示文稿中的数据一致性',
      '提交前检查长标题、换行和最终页面排版',
    ],
  },
];

export const COMPLEX_SOURCES: Record<HostKind, string> = {
  word: reportToText(COMPLEX_REPORT),
  excel: tableToText(COMPLEX_TABLE),
  powerpoint: COMPLEX_SLIDES.map(
    (slide, i) => `${i + 1}. ${slide.title}\n${slide.bullets.join('\n')}`,
  ).join('\n\n'),
};
export const COMPLEX_DESCRIPTIONS: Record<HostKind, string> = {
  word: '6 节长报告、2 张对比表与待办清单。选区演示以文本呈现。',
  excel: '48 行运营数据、8 列、96 个公式，含空值与多行备注。',
  powerpoint: '12 页验收汇报，包含多层次内容与完整行动清单。',
};
