import { parseFormula, parseOutline } from '../core/output';
import type { HostBridge, HostKind, TaskSpec } from '../core/types';
import { parseReport, parseTable, reportToText, tableToText } from '../core/structured-output';
import {
  COMPLEX_REPORT,
  COMPLEX_TABLE,
  COMPLEX_SLIDES,
  COMPLEX_SOURCES,
  type DemoScenario,
} from './complex-samples';

export const DEMO_SOURCES: Record<HostKind, string> = {
  word: '本季度，我们的团队在三个方面取得了进展。首先，客户反馈的平均响应时间从 24 小时缩短至 8 小时。其次，新版入门指南已上线，帮助新用户更快完成配置。最后，我们计划在下季度进一步优化跨团队协作流程，让信息传递更加顺畅。',
  excel: JSON.stringify(
    {
      address: '销售汇总!A1:D5',
      values: [
        ['月份', '华东', '华南', '华北'],
        ['1月', 128, 96, 82],
        ['2月', 156, 112, 90],
        ['3月', 184, 138, 108],
        ['合计', 468, 346, 280],
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
  powerpoint: '季度工作回顾\n客户响应效率提升\n新版入门指南上线\n下季度优化跨团队协作流程',
};

export function createDemoBridge(
  host: HostKind,
  onWrite: (text: string) => void,
  scenario: DemoScenario = 'basic',
): HostBridge {
  let content = scenario === 'complex' ? COMPLEX_SOURCES[host] : DEMO_SOURCES[host];
  return {
    host,
    mode: 'demo',
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
      const original = content;
      return {
        text: original,
        label:
          scenario === 'complex'
            ? '演示 · 复杂文档选区'
            : host === 'excel'
              ? '演示 · 销售汇总!A1:D5'
              : '演示 · 示例选区',
        replace: async (text) => {
          if (content !== original) throw new Error('示例内容已经变化，请重新读取。');
          content = text;
          onWrite(content);
        },
      };
    },
    prepareWrite: async (kind, source, output) => {
      if (kind === 'replace' && source?.replace)
        return {
          label: '替换演示选区',
          before: source.text,
          commit: source.replace,
        };
      return {
        label:
          kind === 'table'
            ? `新建演示工作表「${parseTable(output || '').name}」`
            : kind === 'document'
              ? '在演示文档末尾插入排版报告'
              : kind === 'formula'
                ? '写入演示单元格 B6'
                : kind === 'slides'
                  ? '添加到演示幻灯片'
                  : '插入演示文档末尾',
        before: kind === 'formula' ? '（空单元格）' : '仅修改此页面中的模拟文档',
        commit: async (text) => {
          if (kind === 'formula') parseFormula(text);
          if (kind === 'slides') parseOutline(text);
          const display =
            kind === 'document'
              ? reportToText(parseReport(text))
              : kind === 'table'
                ? tableToText(parseTable(text))
                : text;
          content = `${content}\n\n${display}`;
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
      reject(new Error('已取消生成。'));
      return;
    }
    const abort = () => {
      clearTimeout(timer);
      reject(new Error('已取消生成。'));
    };
    const timer = setTimeout(() => {
      signal.removeEventListener('abort', abort);
      resolve();
    }, 650);
    signal.addEventListener('abort', abort, { once: true });
  });
  if (task.output === 'document') return JSON.stringify(COMPLEX_REPORT, null, 2);
  if (task.output === 'table') return JSON.stringify(COMPLEX_TABLE, null, 2);
  if (scenario === 'complex') {
    if (task.output === 'outline') return JSON.stringify({ slides: COMPLEX_SLIDES }, null, 2);
    if (task.output === 'formula')
      return "=IFERROR(SUMIFS('区域运营演示'!$F$2:$F$49,'区域运营演示'!$B$2:$B$49,\"华东\")/SUMIFS('区域运营演示'!$D$2:$D$49,'区域运营演示'!$B$2:$B$49,\"华东\"),0)";
    if (task.id === 'summarize' || task.id === 'organize')
      return '演示验收摘要\n\n48 项检查中有 42 项通过，通过率为 87.5%。其余 4 项待复核，2 项待外部资料。\n\n下一步先补齐资料和验收证据，再复算指标、更新结论。此结果为固定合成样例。';
    if (task.id === 'analyze' || task.id === 'ask')
      return '合成数据检查\n\n表格包含 12 个月、4 个区域，共 48 行。差额 = 收入 − 成本；差额率 = 差额 ÷ 收入。\n\n首行收入 1200，成本 800，差额 400，差额率约 33.3%。不要把百分比直接相加；比较总体比例时应以收入加权。\n\n备注包含空值和换行，不代表数据丢失。这是固定演示分析，未调用 AI。';
    if (task.id === 'explain')
      return '演示公式使用 SUMIFS 按区域汇总差额与收入，再计算比率。IFERROR 在分母为零或引用发生错误时返回 0；正式分析应区分真实零值与被掩盖的错误。\n\n此说明为固定演示内容。';
    if (task.id === 'rewrite')
      return `复杂文档演示预览\n\n${source}\n\n以上保留合成样稿内容，用于验证长文编辑与写回，未调用 AI。`;
  }
  if (task.output === 'formula') return '=SUM(B2:B4)';
  if (task.output === 'outline')
    return JSON.stringify(
      {
        slides: [
          {
            title: '季度工作回顾',
            bullets: [
              '围绕客户体验，回顾本季度重点工作',
              '以下为固定演示样例，请连接模型生成真实内容',
            ],
          },
          {
            title: '客户响应效率提升',
            bullets: ['平均响应时间从 24 小时缩短至 8 小时', '持续收集反馈，改进服务流程'],
          },
          {
            title: '让入门更轻松',
            bullets: ['新版入门指南已上线', '帮助新用户更快完成配置'],
          },
          {
            title: '下一步行动',
            bullets: ['优化跨团队协作', '明确交接节点与反馈机制'],
          },
        ],
      },
      null,
      2,
    );
  if (task.id === 'explain')
    return '示例公式 =SUM(B2:B4) 会将 B2、B3、B4 三个单元格的数值相加。\n\n在演示数据中：128 + 156 + 184 = 468。\n空白单元格和文本会被忽略；引用范围中存在错误值时，公式也可能返回错误。';
  if (task.id === 'ask' || task.id === 'analyze')
    return '演示数据观察\n\n• 三个区域的月度数值均连续增长。\n• 华东季度合计 468，高于华南的 346 和华北的 280。\n• 下一步可以补充成本和目标数据，进一步比较各区域的完成情况。\n\n这些是固定演示数据的分析，真实问题请连接你的模型。';
  if (task.id === 'translate')
    return 'This quarter, our team made progress in three areas. The average response time to customer feedback decreased from 24 hours to 8 hours. A new onboarding guide was launched to help users get started faster. Next quarter, we plan to improve cross-team collaboration and make information sharing smoother.';
  if (task.id === 'summarize' || task.id === 'organize')
    return '本季度工作重点\n\n• 客户反馈响应时间从 24 小时缩短至 8 小时。\n• 新版入门指南已上线，简化用户配置流程。\n• 下季度将重点优化跨团队协作。';
  if (task.id === 'generate')
    return `关于「${instruction.slice(0, 80)}」的示例草稿\n\n项目背景\n请在这里补充目标受众、背景信息和预期成果。\n\n工作计划\n1. 明确目标与交付范围。\n2. 整理关键资料并完成初稿。\n3. 收集反馈，确认下一步行动。\n\n此处为演示模板。配置真实模型后，将根据你的要求生成内容。`;
  if (source === DEMO_SOURCES.word)
    return '本季度，团队在客户服务、用户引导和协作规划方面取得了进展：\n\n客户反馈的平均响应时间由 24 小时缩短至 8 小时；新版入门指南正式上线，帮助新用户更快完成配置。\n\n下季度，我们将继续优化跨团队协作流程，让信息流转更高效。';
  return `演示预览\n\n${source}\n\n这是交互演示，未调用 AI。连接真实模型后，将根据你的内容与要求完成改写。`;
}
