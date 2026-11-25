import { getUiLanguage, type UiLanguage } from '../i18n';
import type { HostKind, ReportContent, SlideContent, TableContent } from '../core/types';
import { reportToText, tableToText } from '../core/structured-output';

export type DemoScenario = 'basic' | 'complex';

// Deterministic, synthetic fixtures shared by the browser demo and regression tests.
function makeComplexSamples(language: UiLanguage) {
  const text = (zh: string, en: string) => (language === 'en' ? en : zh);
  const COMPLEX_REPORT: ReportContent = {
    title: text('跨团队项目验收报告', 'Cross-team project acceptance report'),
    sections: [
      {
        heading: text('验收结论与范围', 'Scope and findings'),
        paragraphs: [
          text(
            '本报告为合成演示样稿。项目分为需求梳理、数据迁移、使用培训和交付验收四个阶段。当前已完成 48 项检查，其中 42 项通过、4 项需要复核、2 项等待外部资料。未关闭事项不计入已通过项目。',
            'This fictional project covers requirements, data migration, training, and acceptance. Of 48 checks, 42 passed, four need review, and two await external documents. Open items do not count as passed.',
          ),
          text(
            '本轮验收覆盖三个区域、四个交付小组和两类使用场景。附件、流程图与原始数据应在正式交付前逐项核对，不能用摘要代替附件检查。',
            'The review covers three regions, four delivery teams, and two use cases. Check the attachments, process diagrams, and raw data before delivery.',
          ),
        ],
        bullets: [
          text('通过率按 42 ÷ 48 计算，为 87.5%。', 'The pass rate is 42 ÷ 48, or 87.5%.'),
          text(
            '上线决定需以全部必需事项关闭后的复核结果为准。',
            'Review all required items before deciding whether to release.',
          ),
        ],
      },
      {
        heading: text('指标口径与结果', 'Measures and results'),
        paragraphs: [
          text(
            '所有数量均为演示数据。响应时间使用小时，完成率使用通过检查数与总检查数的比值；平均值不能代替每个小组的达标情况。',
            'All figures are synthetic. Response times are in hours. The pass rate is passed checks divided by all checks. An average cannot show whether each team met its target.',
          ),
        ],
        bullets: [],
        table: {
          headers: [
            text('指标', 'Measure'),
            text('期初', 'Baseline'),
            text('本期', 'Current'),
            text('复核说明', 'Review note'),
          ],
          rows: [
            [
              text('平均响应时间', 'Average response time'),
              text('24 小时', '24 hours'),
              text('8 小时', '8 hours'),
              text('按相同时间窗口比较', 'Compare equivalent time windows'),
            ],
            [
              text('验收检查', 'Acceptance checks'),
              text('48 项', '48 checks'),
              text('42 项通过', '42 passed'),
              text('4 项复核，2 项待资料', '4 under review, 2 awaiting documents'),
            ],
            [
              text('通过率', 'Pass rate'),
              text('待补充', 'TBD'),
              '87.5%',
              text('分母包含未关闭事项', 'The denominator includes open items'),
            ],
          ],
        },
      },
      {
        heading: text('数据迁移与一致性', 'Data migration and consistency'),
        paragraphs: Array.from({ length: 4 }, (_, i) =>
          text(
            `第 ${i + 1} 批数据需要核对记录标识、字段类型、关联关系与时间范围。保留空白和零值之间的区别，百分比应按原始数值计算后再格式化显示。带换行、制表符、中文、英文或特殊符号的单元格必须完整保留；如出现重复记录，应先确认业务含义再决定是否合并。此段为合成的长段落，用于检查内容阅读与编辑。`,
            `Batch ${i + 1}: verify record IDs, field types, relationships, and date ranges. Distinguish blanks from zero and calculate percentages from raw values. Preserve cells containing newlines, tabs, multiple languages, or special characters. Review the meaning of duplicate records before merging them. This fictional passage tests how a longer section reads and edits.`,
          ),
        ),
        bullets: [
          text(
            '抽样不能替代关键字段的完整校验。',
            'Check all critical fields as well as the sample.',
          ),
          text(
            '缺失资料单独登记，负责人填写后再次复核。',
            'Track missing materials separately and review them when the owner supplies them.',
          ),
        ],
      },
      {
        heading: text('风险与待办', 'Risks and outstanding actions'),
        paragraphs: [
          text(
            '以下事项尚未关闭，状态变化需要负责人确认。日期与负责人保留为占位信息，避免演示数据被误认为实际承诺。',
            'These items remain open. Owners must confirm status changes. Dates and owners are placeholders in this example.',
          ),
        ],
        bullets: [],
        table: {
          headers: [
            text('事项', 'Item'),
            text('负责人', 'Owner'),
            text('状态', 'Status'),
            text('完成条件', 'Completion criteria'),
          ],
          rows: [
            [
              text('4 项复核', '4 checks to review'),
              text('待补充', 'TBD'),
              text('待复核', 'Under review'),
              text('复核通过并保存证据', 'Pass the review and retain the evidence'),
            ],
            [
              text('2 项外部资料', '2 external documents'),
              text('待补充', 'TBD'),
              text('待提供', 'Awaiting documents'),
              text('资料到齐且内容一致', 'All documents received and consistent'),
            ],
            [
              text('使用培训', 'User training'),
              text('待补充', 'TBD'),
              text('待确认', 'Awaiting confirmation'),
              text('记录参加情况和遗留问题', 'Record attendance and open questions'),
            ],
          ],
        },
      },
      {
        heading: text('跨团队交接安排', 'Cross-team handover'),
        paragraphs: [
          text(
            '每个交付小组整理本组输入、交付物、依赖事项和验收记录。接收方核对文件版本与内容后确认接收；发现差异时保留原始记录，补齐说明后再更新。',
            'Each delivery team records its inputs, deliverables, dependencies, and acceptance evidence. The receiving team verifies file versions and content. Retain the original records and document discrepancies before updating them.',
          ),
          text(
            '对跨区域协作，明确时间口径与责任边界。工作表中的公式引用需包含准确的区域和期间，不能将多期累计值重复相加。',
            'For work across regions, agree on reporting periods and ownership. Formulas should reference the correct regions and periods. Avoid adding cumulative totals twice.',
          ),
        ],
        bullets: [
          text(
            '交接材料包含交付清单与未关闭事项。',
            'Include the delivery checklist and open items in the handover.',
          ),
          text(
            '接收记录与补充资料保持可追溯。',
            'Keep a record of receipt and any supporting material.',
          ),
        ],
      },
      {
        heading: text('后续行动与复核', 'Next steps and follow-up'),
        paragraphs: [
          text(
            '先关闭待复核事项，再更新验收结论。正式版应补齐负责人、时间和证据链接；本演示不包含真实个人信息或客户数据。',
            'Close items under review before updating the conclusion. Add owners, dates, and links to evidence in the final version. This example contains synthetic data.',
          ),
        ],
        bullets: [
          text(
            '逐项补齐资料和验收证据。',
            'Complete the supporting documents and acceptance evidence.',
          ),
          text('复算关键指标并检查格式。', 'Recalculate key metrics and check formatting.'),
          text(
            '保留原始文件，在确认后提交新版本。',
            'Keep the original and submit the new version after approval.',
          ),
        ],
      },
    ],
  };

  const COMPLEX_TABLE: TableContent = {
    name: text('区域运营演示', 'Regional Operations Demo'),
    headers: [
      text('月份', 'Month'),
      text('区域', 'Region'),
      text('项目', 'Project'),
      text('收入', 'Revenue'),
      text('成本', 'Cost'),
      text('差额', 'Margin'),
      text('差额率', 'Margin rate'),
      text('备注', 'Notes'),
    ],
    formats: ['text', 'text', 'text', 'number', 'number', 'number', 'percent', 'text'],
    rows: Array.from({ length: 48 }, (_, i) => {
      const row = i + 2;
      return [
        text(`${Math.floor(i / 4) + 1}月`, `Month ${Math.floor(i / 4) + 1}`),
        [
          text('华东', 'East'),
          text('华南', 'South'),
          text('华北', 'North'),
          text('西南', 'Southwest'),
        ][i % 4],
        text(`项目 ${String(i + 1).padStart(3, '0')}`, `Project ${String(i + 1).padStart(3, '0')}`),
        1200 + i * 37.25,
        800 + i * 19.125,
        { formula: `=D${row}-E${row}` },
        { formula: `=IFERROR(F${row}/D${row},0)` },
        i === 0
          ? text('合成数据\n保留两行备注', 'Synthetic data\nKeep this second line')
          : i === 47
            ? text('末行核对', 'Verify final row')
            : null,
      ];
    }),
  };

  const COMPLEX_SLIDES: SlideContent[] = [
    {
      title: text('项目验收与下一阶段安排', 'Project acceptance and next steps'),
      bullets: [
        text('合成演示数据：48 项检查，42 项通过', 'Demo data: 48 checks, 42 passed'),
        text(
          '4 项需要复核，2 项等待外部资料',
          'Four checks need review and two await outside documents',
        ),
        text(
          '重点说明验收口径、未关闭事项与责任交接',
          'Explain the criteria, open items, and handover responsibilities',
        ),
      ],
    },
    ...[
      text('验收范围', 'Acceptance scope'),
      text('需求梳理', 'Requirements review'),
      text('数据迁移', 'Data migration'),
      text('数值与公式校验', 'Values and formula checks'),
      text('区域协作', 'Regional collaboration'),
      text('培训与使用反馈', 'Training and user feedback'),
      text('交付材料', 'Delivery materials'),
      text('风险与依赖', 'Risks and dependencies'),
      text('后续行动', 'Next steps'),
      text('复核机制', 'Review process'),
    ].map((title) => ({
      title,
      bullets: [
        text(
          '以原始资料为依据，保留事实、数字和单位',
          'Use source records and preserve facts, figures, and units',
        ),
        text(
          '区分已完成事项与等待补充的信息',
          'Distinguish completed items from missing information',
        ),
        text(
          '明确负责人、完成条件与复核时间',
          'Identify owners, completion criteria, and review dates',
        ),
        text(
          '缺失信息用“待补充”标示，不代替实际确认',
          'Mark missing details as TBD until confirmed',
        ),
      ],
    })),
    {
      title: text('验收确认与行动清单', 'Acceptance and action list'),
      bullets: [
        text(
          '关闭全部必需事项后更新结论',
          'Update the conclusion when all required items are closed',
        ),
        text(
          '核对文档、表格及演示文稿中的数据一致性',
          'Check consistency across the document, workbook, and slides',
        ),
        text(
          '提交前检查长标题、换行和最终页面排版',
          'Inspect long titles, wrapping, and final slide layout',
        ),
      ],
    },
  ];

  const COMPLEX_SOURCES: Record<HostKind, string> = {
    word: reportToText(COMPLEX_REPORT),
    excel: tableToText(COMPLEX_TABLE),
    powerpoint: COMPLEX_SLIDES.map(
      (slide, i) => `${i + 1}. ${slide.title}\n${slide.bullets.join('\n')}`,
    ).join('\n\n'),
  };
  const COMPLEX_DESCRIPTIONS: Record<HostKind, string> = {
    word: text(
      '6 节长报告、2 张对比表与待办清单。选区演示以文本呈现。',
      'A six-section report with two comparison tables and an action list. The selection appears as text.',
    ),
    excel: text(
      '48 行运营数据、8 列、96 个公式，含空值与多行备注。',
      '48 data rows, eight columns, and 96 formulas, including blanks and multiline notes.',
    ),
    powerpoint: text(
      '12 页验收汇报，包含多层次内容与完整行动清单。',
      'A 12-slide review with section details and a complete action list.',
    ),
  };

  return {
    report: COMPLEX_REPORT,
    table: COMPLEX_TABLE,
    slides: COMPLEX_SLIDES,
    sources: COMPLEX_SOURCES,
    descriptions: COMPLEX_DESCRIPTIONS,
  };
}
const chineseSamples = makeComplexSamples('zh-CN');
let englishSamples: ReturnType<typeof makeComplexSamples> | undefined;
export function getComplexSamples(language = getUiLanguage()) {
  return language === 'en' ? (englishSamples ??= makeComplexSamples('en')) : chineseSamples;
}
export const {
  report: COMPLEX_REPORT,
  table: COMPLEX_TABLE,
  slides: COMPLEX_SLIDES,
  sources: COMPLEX_SOURCES,
  descriptions: COMPLEX_DESCRIPTIONS,
} = chineseSamples;
