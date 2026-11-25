import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { mkdtemp, rm, writeFile, copyFile, rename, unlink } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { constants } from 'node:fs';
import { reportRtf } from './report.mjs';
import { wordMarkdownFormat } from './word-markdown.mjs';
const hosts = {
  word: ['replace', 'insert', 'document'],
  excel: ['formula', 'table'],
  powerpoint: ['slides'],
};

export function runOffice(root, input) {
  return new Promise((resolve, reject) => {
    const mac = process.platform === 'darwin';
    const binary = mac ? '/usr/bin/osascript' : join(root, 'native', 'Office.exe');
    const script = join(root, 'native', 'office-mac.js');
    const args = mac ? ['-l', 'JavaScript', script] : [];
    const child = spawn(binary, args, { windowsHide: true, stdio: ['pipe', 'pipe', 'ignore'] });
    const chunks = [];
    let length = 0;
    const timer = setTimeout(() => child.kill(), 95000);
    child.once('error', (e) => {
      clearTimeout(timer);
      reject(e);
    });
    child.stdout.on('data', (chunk) => {
      length += chunk.length;
      if (length > 6000000) child.kill();
      else chunks.push(chunk);
    });
    child.stdin.on('error', () => {});
    child.once('close', (code) => {
      clearTimeout(timer);
      try {
        if (code !== 0)
          throw Error('Office 未响应。请在目标应用打开文档，并允许系统请求的自动化访问。');
        const response = JSON.parse(
          Buffer.concat(chunks)
            .toString('utf8')
            .replace(/^\uFEFF/, ''),
        );
        if (!response.ok) throw Error(response.message || 'Office 未完成操作。');
        resolve(response.value);
      } catch (e) {
        reject(e);
      }
    });
    child.stdin.end(JSON.stringify(input), 'utf8');
  });
}

export function createOfficeSession(host, root, { run = runOffice, saveDeck, openDeck } = {}) {
  const targets = new Map();
  let busy = false,
    closed = false;
  function remember(entry) {
    while (targets.size >= 32) targets.delete(targets.keys().next().value);
    const token = randomUUID();
    targets.set(token, entry);
    return token;
  }
  async function dispatch(action, data = {}) {
    if (closed) throw Error('工作台已关闭。');
    if (action === 'release') {
      targets.delete(data.token);
      return;
    }
    if (busy) throw Error('上一项 Office 操作仍在进行。');
    busy = true;
    let temp;
    try {
      const office = (op, body = {}) => run(root, { host, op, ...body });
      if (action === 'selection' || action === 'target') {
        const kind = action === 'selection' ? 'replace' : data.kind;
        if (action === 'target' && !hosts[host].includes(kind))
          throw Error('当前应用不支持这项写入。');
        const { guard, ...view } = await office(action, { kind });
        const token = remember({ guard, kind, canWrite: action === 'target' || view.writable });
        return { ...view, token };
      }
      if (
        action === 'presentation' &&
        host === 'powerpoint' &&
        ['all', 'selected'].includes(data.scope)
      ) {
        const { guard, snapshot } = await office(action, { scope: data.scope });
        return { token: remember({ guard, snapshot, kind: 'presentation' }), snapshot };
      }
      if (action === 'presentationTarget' && host === 'powerpoint') {
        const { guard, target } = await office(action);
        return { token: remember({ guard, target, kind: 'presentation-insert' }), target };
      }
      const entry = targets.get(data.token);
      if (action !== 'create' && !entry) throw Error('写入位置已过期，请重新读取。');
      if (action === 'write' && entry.canWrite) {
        const wordText = host === 'word' && ['replace', 'insert'].includes(entry.kind);
        if (
          data.value == null ||
          (wordText ? typeof data.value !== 'string' : JSON.stringify(data.value).length > 2000000)
        )
          throw Error('写入内容无效或过大。');
        const body = { kind: entry.kind, guard: entry.guard, value: data.value };
        const formatting = wordText ? wordMarkdownFormat(data.value) : null;
        if (formatting) {
          temp = await mkdtemp(join(tmpdir(), 'triobridge-word-'));
          body.file = join(temp, randomUUID() + '.htm');
          body.block = formatting.block;
          await writeFile(
            body.file,
            '\uFEFF<!doctype html><html><head><meta charset="utf-8"></head><body>' +
              formatting.html +
              '</body></html>',
            { mode: 0o600 },
          );
          delete body.value;
        } else if (entry.kind === 'document') {
          temp = await mkdtemp(join(tmpdir(), 'triobridge-report-'));
          body.file = join(temp, 'report.rtf');
          await writeFile(body.file, reportRtf(data.value), { mode: 0o600 });
          delete body.value;
        }
        const value = await office('write', body);
        targets.delete(data.token);
        return value;
      }
      if (host === 'powerpoint' && action === 'edit' && entry.kind === 'presentation') {
        const { proposal } = data;
        if (!proposal || !Array.isArray(proposal.changes) || proposal.changes.length > 600)
          throw Error('幻灯片改动无效。');
        const value = await office('edit', {
          guard: entry.guard,
          snapshot: entry.snapshot,
          proposal,
        });
        targets.delete(data.token);
        return value;
      }
      if (
        host === 'powerpoint' &&
        (action === 'create' || (action === 'append' && entry.kind === 'presentation-insert'))
      ) {
        if (action === 'append' && (!Number.isSafeInteger(data.count) || data.count < 1))
          throw Error('待插入页数无效。');
        if (typeof data.base64 !== 'string' || data.base64.length > 90000000)
          throw Error('演示稿文件过大。');
        const bytes = Buffer.from(data.base64, 'base64');
        if (bytes.length < 22 || bytes.readUInt32LE(0) !== 0x04034b50)
          throw Error('演示稿格式无效。');
        temp = await mkdtemp(join(tmpdir(), 'triobridge-deck-'));
        const file = join(temp, 'deck.pptx');
        await writeFile(file, bytes, { mode: 0o600 });
        if (action === 'create') {
          const destination = await saveDeck();
          if (!destination) throw Error('已取消保存演示稿。');
          const pending = join(dirname(destination), '.triobridge-' + randomUUID() + '.pptx');
          try {
            await copyFile(file, pending, constants.COPYFILE_EXCL);
            await rename(pending, destination);
          } finally {
            await unlink(pending).catch(() => {});
          }
          try {
            if (openDeck) await openDeck(destination);
            else await office('openDeck', { file: destination });
          } catch {
            throw Error('演示稿已保存。请在 PowerPoint 中打开刚保存的文件。');
          }
          return;
        }
        await office('append', {
          file,
          guard: entry.guard,
          target: entry.target,
          count: data.count,
        });
        targets.delete(data.token);
        return;
      }
      throw Error('此操作不适用于当前文档。');
    } finally {
      if (temp) await rm(temp, { recursive: true, force: true });
      busy = false;
    }
  }
  return {
    close() {
      closed = true;
      targets.clear();
    },
    async invoke(action, data) {
      try {
        return { ok: true, value: await dispatch(action, data) };
      } catch (e) {
        return { ok: false, message: String(e.message).slice(0, 360) };
      }
    },
  };
}
