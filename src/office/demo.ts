import { parseFormula, parseOutline } from '../core/output';
import type { HostBridge, HostKind, TaskSpec } from '../core/types';

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

export function createDemoBridge(host: HostKind, onWrite: (text: string) => void): HostBridge {
  let content = DEMO_SOURCES[host];
  return {
    host,
    mode: 'demo',
    capabilities: {
      read: true,
      replace: host !== 'excel',
      insert: host === 'word',
      formula: host === 'excel',
      slides: host === 'powerpoint',
    },
    readSelection: async () => {
      const original = content;
      return {
        text: original,
        label: host === 'excel' ? '演示 · 销售汇总!A1:D5' : '演示 · 示例选区',
        replace: async (text) => {
          if (content !== original) throw new Error('示例内容已经变化，请重新读取。');
          content = text;
          onWrite(content);
        },
      };
    },
    prepareWrite: async (kind, source) => {
      if (kind === 'replace' && source?.replace)
        return {
          label: '替换演示选区',
          before: source.text,
          commit: source.replace,
        };
      return {
        label:
          kind === 'formula'
            ? '写入演示单元格 B6'
            : kind === 'slides'
              ? '添加到演示幻灯片'
              : '插入演示文档末尾',
        before: kind === 'formula' ? '（空单元格）' : '仅修改此页面中的模拟文档',
        commit: async (text) => {
          if (kind === 'formula') parseFormula(text);
          if (kind === 'slides') parseOutline(text);
          content = `${content}\n\n${text}`;
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
