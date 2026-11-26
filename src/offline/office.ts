import { uiMessage } from '../i18n-errors';
import type { HostBridge, SelectionSnapshot, WriteKind } from '../core/types';
import { parseFormula, parseOutline } from '../core/output';
import { parseReport, parseTable, tableMatrix, NUMBER_FORMATS } from '../core/structured-output';
import { bytes64 } from '../presentation/package-xml';
import { validateEdit } from '../presentation/editing';
import type { PresentationInsertTarget, PresentationSnapshot } from '../presentation/types';
import type { NativeWorkbench } from './transport';

interface Target {
  token: string;
  label: string;
  text: string;
  writable: boolean;
  warning?: string;
}
export function createNativeBridge(native: NativeWorkbench): HostBridge {
  const host = native.host;
  const word = host === 'word',
    excel = host === 'excel',
    ppt = host === 'powerpoint';
  const slideTokens = new WeakMap<PresentationSnapshot, string>();
  const insertTokens = new WeakMap<PresentationInsertTarget, string>();
  async function call<T>(action: string, data?: unknown): Promise<T> {
    const result = await native.office(action, data);
    if (!result.ok) throw new Error(result.message || 'Office 未完成操作，请检查当前文档。');
    return result.value as T;
  }
  async function release(token: string) {
    await call('release', { token });
  }
  function prepareValue(kind: WriteKind, text: string) {
    if (kind === 'document') return parseReport(text);
    if (kind === 'table') {
      const table = parseTable(text);
      return {
        name: table.name,
        matrix: tableMatrix(table),
        formats: table.formats.map((f) => NUMBER_FORMATS[f]),
      };
    }
    if (kind === 'formula') return parseFormula(text);
    if (kind === 'slides') return parseOutline(text);
    return text;
  }
  return {
    host,
    mode: 'office',
    capabilities: {
      read: true,
      replace: word,
      insert: word,
      document: word,
      formula: excel,
      table: excel,
      slides: ppt,
    },
    async readSelection(): Promise<SelectionSnapshot> {
      const target = await call<Target>('selection');
      return {
        text: target.text,
        label: target.label,
        warning: target.warning,
        ...(target.writable && word
          ? {
              replace: async (text: string) => {
                await call('write', { token: target.token, value: text });
              },
            }
          : {}),
        dispose: () => release(target.token),
      };
    },
    async prepareWrite(kind, source, output) {
      if (kind === 'replace' && source?.replace)
        return {
          label: source.label,
          before: source.text,
          warning: '替换会改变选区内的文字和格式，请核对预览。',
          commit: source.replace,
        };
      prepareValue(kind, output || '');
      const target = await call<Target>('target', { kind });
      return {
        label: target.label,
        before: kind === 'table' || kind === 'slides' ? uiMessage(target.text) : target.text,
        warning: target.warning,
        dispose: () => release(target.token),
        commit: async (text) => {
          await call('write', { token: target.token, value: prepareValue(kind, text) });
        },
      };
    },
    presentation: ppt
      ? {
          capabilities: {
            read: true,
            selected: true,
            reorder: navigator.platform.startsWith('Win'),
            insert: navigator.platform.startsWith('Win'),
            create: true,
            thumbnail: false,
            backup: true,
            backgrounds: true,
            gradient: true,
          },
          async read(scope) {
            const result = await call<{ token: string; snapshot: PresentationSnapshot }>(
              'presentation',
              { scope },
            );
            slideTokens.set(result.snapshot, result.token);
            return result.snapshot;
          },
          async apply(snapshot, proposal) {
            const checked = validateEdit(proposal, snapshot);
            await call('edit', { token: slideTokens.get(snapshot), proposal: checked });
          },
          async prepareInsert() {
            const result = await call<{ token: string; target: PresentationInsertTarget }>(
              'presentationTarget',
            );
            insertTokens.set(result.target, result.token);
            return result.target;
          },
          async insert(base64, count, target, afterSlideId) {
            await call('append', {
              token: insertTokens.get(target),
              base64,
              count,
              ...(afterSlideId ? { afterSlideId } : {}),
            });
          },
          async create(base64) {
            await call('create', { base64 });
          },
          async thumbnail() {
            throw Error('此入口无法预览页面缩略图。');
          },
          async backup() {
            return new Blob([bytes64(await call<string>('presentationCopy'))], {
              type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            });
          },
        }
      : undefined,
  };
}
