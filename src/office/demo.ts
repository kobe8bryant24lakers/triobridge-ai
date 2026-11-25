import { getUiLanguage, t, type UiLanguage } from '../i18n';
import { parseFormula, parseOutline } from '../core/output';
import type { HostBridge, HostKind, TaskSpec } from '../core/types';
import { parseReport, parseTable, reportToText, tableToText } from '../core/structured-output';
import { createDemoPresentation } from './presentation-demo';
import { getComplexSamples, type DemoScenario } from './complex-samples';

function makeDemoSources(language: UiLanguage): Record<HostKind, string> {
  const text = (zh: string, en: string) => (language === 'en' ? en : zh);
  return {
    word: text(
      '本季度，我们的团队在三个方面取得了进展。首先，客户反馈的平均响应时间从 24 小时缩短至 8 小时。其次，新版入门指南已上线，帮助新用户更快完成配置。最后，我们计划在下季度进一步优化跨团队协作流程，让信息传递更加顺畅。',
      'This quarter, our team made progress in three areas. Average response time for customer feedback fell from 24 hours to 8 hours. The new onboarding guide helps users finish setup faster. Next quarter, we plan to improve cross-team handovers.',
    ),
    excel: JSON.stringify(
      {
        address: text('销售汇总!A1:D5', 'Sales Summary!A1:D5'),
        values: [
          [
            text('月份', 'Month'),
            text('华东', 'East'),
            text('华南', 'South'),
            text('华北', 'North'),
          ],
          [text('1月', 'January'), 128, 96, 82],
          [text('2月', 'February'), 156, 112, 90],
          [text('3月', 'March'), 184, 138, 108],
          [text('合计', 'Total'), 468, 346, 280],
        ],
        formulas: [
          ['', '', '', ''],
          ['', '', '', ''],
          ['', '', '', ''],
          ['', '', '', ''],
          ['', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)'],
        ],
      },
      null,
      2,
    ),
    powerpoint: text(
      '季度工作回顾\n客户响应效率提升\n新版入门指南上线\n下季度优化跨团队协作流程',
      'Quarterly review\nFaster customer responses\nNew onboarding guide released\nImprove cross-team handovers next quarter',
    ),
  };
}
export const DEMO_SOURCES = makeDemoSources('zh-CN');
const englishSources = makeDemoSources('en');
function demoSources() {
  return getUiLanguage() === 'en' ? englishSources : DEMO_SOURCES;
}

export function createDemoBridge(
  host: HostKind,
  onWrite: (text: string) => void,
  scenario: DemoScenario = 'basic',
): HostBridge {
  let content: string | undefined;
  let sheetAddress: string;
  function documentContent() {
    if (content === undefined) {
      sheetAddress = t('销售汇总!A1:D5', 'Sales Summary!A1:D5');
      content = scenario === 'complex' ? getComplexSamples().sources[host] : demoSources()[host];
    }
    return content;
  }

  return {
    host,
    mode: 'demo',
    ...(host === 'powerpoint' ? { presentation: createDemoPresentation(onWrite) } : {}),
    capabilities: {
      read: true,
      replace: host !== 'excel',
      insert: host === 'word',
      formula: host === 'excel',
      slides: host === 'powerpoint',
      document: host === 'word',
      table: host === 'excel',
    },
    readSelection: async () => {
      const original = documentContent();
      return {
        text: original,
        get label() {
          return scenario === 'complex'
            ? t('演示 · 复杂文档选区', 'Demo · Complex sample selection')
            : host === 'excel'
              ? t(`演示 · ${sheetAddress}`, `Demo · ${sheetAddress}`)
              : t('演示 · 示例选区', 'Demo · Sample selection');
        },
        replace: async (text) => {
          if (content !== original)
            throw new Error(
              t('示例内容已经变化，请重新读取。', 'The sample content has changed. Read it again.'),
            );
          content = text;
          onWrite(content);
        },
      };
    },
    prepareWrite: async (kind, source, output) => {
      if (kind === 'replace' && source?.replace)
        return {
          label: t('替换演示选区', 'Replace the sample selection'),
          before: source.text,
          commit: source.replace,
        };
      const tableName = kind === 'table' ? parseTable(output || '').name : '';
      return {
        label:
          kind === 'table'
            ? t(`新建演示工作表「${tableName}」`, `Create sample worksheet “${tableName}”`)
            : kind === 'document'
              ? t(
                  '在演示文档末尾插入排版报告',
                  'Insert a formatted report at the end of the demo document',
                )
              : kind === 'formula'
                ? t('写入演示单元格 B6', 'Write to sample cell B6')
                : kind === 'slides'
                  ? t('添加到演示幻灯片', 'Append to the demo presentation')
                  : t('插入演示文档末尾', 'Insert at the end of the demo document'),
        before:
          kind === 'formula'
            ? t('（空单元格）', '(Empty cell)')
            : t(
                '仅修改此页面中的模拟文档',
                'Only the simulated document on this page will change.',
              ),
        commit: async (text) => {
          if (kind === 'formula') parseFormula(text);
          if (kind === 'slides') parseOutline(text);
          const display =
            kind === 'document'
              ? reportToText(parseReport(text))
              : kind === 'table'
                ? tableToText(parseTable(text))
                : text;
          content = `${documentContent()}\n\n${display}`;
          onWrite(content);
        },
      };
    },
  };
}

export async function demoComplete(
  task: TaskSpec,
  source: string,
  instruction: string,
  signal: AbortSignal,
  scenario: DemoScenario = 'basic',
): Promise<string> {
  await new Promise<void>((resolve, reject) => {
    if (signal.aborted) {
      reject(new Error(t('已取消生成。', 'Generation canceled.')));
      return;
    }
    const abort = () => {
      clearTimeout(timer);
      reject(new Error(t('已取消生成。', 'Generation canceled.')));
    };
    const timer = setTimeout(() => {
      signal.removeEventListener('abort', abort);
      resolve();
    }, 650);
    signal.addEventListener('abort', abort, { once: true });
  });
  const sample = getComplexSamples();
  if (task.output === 'document') return JSON.stringify(sample.report, null, 2);
  if (task.output === 'table') return JSON.stringify(sample.table, null, 2);
  if (scenario === 'complex') {
    if (task.output === 'outline') return JSON.stringify({ slides: sample.slides }, null, 2);
    if (task.output === 'formula')
      return t(
        "=IFERROR(SUMIFS('区域运营演示'!$F$2:$F$49,'区域运营演示'!$B$2:$B$49,\"华东\")/SUMIFS('区域运营演示'!$D$2:$D$49,'区域运营演示'!$B$2:$B$49,\"华东\"),0)",
        "=IFERROR(SUMIFS('Regional Operations Demo'!$F$2:$F$49,'Regional Operations Demo'!$B$2:$B$49,\"East\")/SUMIFS('Regional Operations Demo'!$D$2:$D$49,'Regional Operations Demo'!$B$2:$B$49,\"East\"),0)",
      );
    if (task.id === 'summarize' || task.id === 'organize')
      return t(
        '演示验收摘要\n\n48 项检查中有 42 项通过，通过率为 87.5%。其余 4 项待复核，2 项待外部资料。\n\n下一步先补齐资料和验收证据，再复算指标、更新结论。此结果为固定合成样例。',
        'Acceptance summary\n\nOf 48 checks, 42 passed (87.5%). Four need review and two await external documents.\n\nComplete the evidence and supporting documents, recalculate the metrics, and update the conclusion. This is a fixed sample.',
      );
    if (task.id === 'analyze' || task.id === 'ask')
      return t(
        '合成数据检查\n\n表格包含 12 个月、4 个区域，共 48 行。差额 = 收入 − 成本；差额率 = 差额 ÷ 收入。\n\n首行收入 1200，成本 800，差额 400，差额率约 33.3%。不要把百分比直接相加；比较总体比例时应以收入加权。\n\n备注包含空值和换行，不代表数据丢失。这是固定演示分析，未调用 AI。',
        'Sample data review\n\nThere are 48 rows for 12 months and four regions. Margin = revenue − cost. Margin rate = margin ÷ revenue.\n\nThe first row shows 1,200 in revenue, 800 in cost, and a 400 margin, about 33.3%. Weight overall rates by revenue instead of summing percentages.\n\nThe notes intentionally include blanks and line breaks. This is a fixed demo analysis.',
      );
    if (task.id === 'explain')
      return t(
        '演示公式使用 SUMIFS 按区域汇总差额与收入，再计算比率。IFERROR 在分母为零或引用发生错误时返回 0；正式分析应区分真实零值与被掩盖的错误。\n\n此说明为固定演示内容。',
        'The formula uses SUMIFS to total margin and revenue by region before dividing. IFERROR returns zero if the denominator is zero or a reference fails. In an actual analysis, distinguish genuine zeros from errors.\n\nThis is a fixed demo explanation.',
      );
    if (task.id === 'rewrite')
      return t(
        `复杂文档演示预览\n\n${source}\n\n以上保留合成样稿内容，用于验证长文编辑与写回，未调用 AI。`,
        `Complex document preview

${source}

This sample retains the source for a long-form editing and writing exercise. No model was called.`,
      );
  }
  if (task.output === 'formula') return '=SUM(B2:B4)';
  if (task.output === 'outline')
    return JSON.stringify(
      {
        slides: [
          {
            title: t('季度工作回顾', 'Quarterly review'),
            bullets: [
              t(
                '围绕客户体验，回顾本季度重点工作',
                "Review this quarter's customer experience work",
              ),
              t(
                '以下为固定演示样例，请连接模型生成真实内容',
                'This is fixed demo content. Connect your model to generate a custom result.',
              ),
            ],
          },
          {
            title: t('客户响应效率提升', 'Faster customer responses'),
            bullets: [
              t(
                '平均响应时间从 24 小时缩短至 8 小时',
                'Average response time fell from 24 hours to 8 hours',
              ),
              t('持续收集反馈，改进服务流程', 'Continue collecting feedback and improving service'),
            ],
          },
          {
            title: t('让入门更轻松', 'Easier onboarding'),
            bullets: [
              t('新版入门指南已上线', 'The new onboarding guide is live'),
              t('帮助新用户更快完成配置', 'Help new users finish setup faster'),
            ],
          },
          {
            title: t('下一步行动', 'Next steps'),
            bullets: [
              t('优化跨团队协作', 'Improve cross-team collaboration'),
              t('明确交接节点与反馈机制', 'Define handover points and a feedback process'),
            ],
          },
        ],
      },
      null,
      2,
    );
  if (task.id === 'explain')
    return t(
      '示例公式 =SUM(B2:B4) 会将 B2、B3、B4 三个单元格的数值相加。\n\n在演示数据中：128 + 156 + 184 = 468。\n空白单元格和文本会被忽略；引用范围中存在错误值时，公式也可能返回错误。',
      'The sample formula =SUM(B2:B4) adds the values in B2, B3, and B4.\n\nFor the demo data, 128 + 156 + 184 = 468. Blanks and text are ignored. An error in the range may produce an error in the result.',
    );
  if (task.id === 'ask' || task.id === 'analyze')
    return t(
      '演示数据观察\n\n• 三个区域的月度数值均连续增长。\n• 华东季度合计 468，高于华南的 346 和华北的 280。\n• 下一步可以补充成本和目标数据，进一步比较各区域的完成情况。\n\n这些是固定演示数据的分析，真实问题请连接你的模型。',
      'Sample data findings\n\n• Monthly totals increased in all three regions.\n• East finished at 468, ahead of South at 346 and North at 280.\n• Add cost and target data to compare performance across regions.\n\nThis fixed analysis covers the sample dataset. Connect a model for other questions.',
    );
  if (task.id === 'translate')
    return 'This quarter, our team made progress in three areas. The average response time to customer feedback decreased from 24 hours to 8 hours. A new onboarding guide was launched to help users get started faster. Next quarter, we plan to improve cross-team collaboration and make information sharing smoother.';
  if (task.id === 'summarize' || task.id === 'organize')
    return t(
      '本季度工作重点\n\n• 客户反馈响应时间从 24 小时缩短至 8 小时。\n• 新版入门指南已上线，简化用户配置流程。\n• 下季度将重点优化跨团队协作。',
      "Quarterly highlights\n\n• Average response time fell from 24 hours to 8 hours.\n• The new onboarding guide makes setup easier.\n• Next quarter's focus is cross-team collaboration.",
    );
  if (task.id === 'generate')
    return t(
      `关于「${instruction.slice(0, 80)}」的示例草稿\n\n项目背景\n请在这里补充目标受众、背景信息和预期成果。\n\n工作计划\n1. 明确目标与交付范围。\n2. 整理关键资料并完成初稿。\n3. 收集反馈，确认下一步行动。\n\n此处为演示模板。配置真实模型后，将根据你的要求生成内容。`,
      `Sample draft for “${instruction.slice(0, 80)}”

Project background
Add the audience, context, and expected outcome.

Work plan
1. Agree on the goal and scope.
2. Gather the material and prepare a draft.
3. Collect feedback and confirm the next steps.

This is a demo template. Connect a model to generate a draft based on your requirements.`,
    );
  if (source === DEMO_SOURCES.word || source === englishSources.word)
    return t(
      '本季度，团队在客户服务、用户引导和协作规划方面取得了进展：\n\n客户反馈的平均响应时间由 24 小时缩短至 8 小时；新版入门指南正式上线，帮助新用户更快完成配置。\n\n下季度，我们将继续优化跨团队协作流程，让信息流转更高效。',
      'This quarter, the team improved customer service and onboarding.\n\nAverage response time fell from 24 hours to 8 hours. The new onboarding guide helps users complete setup faster.\n\nNext quarter, we plan to improve handovers so information flows more easily between teams.',
    );
  return t(
    `演示预览\n\n${source}\n\n这是交互演示，未调用 AI。连接真实模型后，将根据你的内容与要求完成改写。`,
    `Demo preview

${source}

This is a fixed interactive demo. Connect a model for a rewrite tailored to your content and instructions.`,
  );
}
