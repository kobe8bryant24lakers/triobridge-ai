import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from 'fflate';
import { t } from '../i18n';
export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main',
  A = 'http://schemas.openxmlformats.org/drawingml/2006/main',
  C = 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  REL = 'http://schemas.openxmlformats.org/package/2006/relationships',
  CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const nodes = (e: Element | Document, ns: string, n: string) =>
  Array.from(e.getElementsByTagNameNS(ns, n));
export const child = (e: Element | undefined, ns: string, name: string) =>
  e && Array.from(e.children).find((c) => c.namespaceURI === ns && c.localName === name);
export const xml = (e: Element | Document) => new XMLSerializer().serializeToString(e);
export function element(
  doc: Document,
  ns: string,
  name: string,
  attrs: Record<string, string | number> = {},
) {
  const e = doc.createElementNS(ns, name);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  return e;
}
export function need(parent: Element, ns: string, name: string, first = false) {
  const old = child(parent, ns, name.split(':').pop()!);
  if (old) return old;
  const e = element(parent.ownerDocument, ns, name);
  parent.insertBefore(e, first ? parent.firstChild : null);
  return e;
}
export function base64(bytes: Uint8Array) {
  const chunks: string[] = [];
  for (let i = 0; i < bytes.length; i += 24576)
    chunks.push(String.fromCharCode(...bytes.subarray(i, i + 24576)));
  return btoa(chunks.join(''));
}
export const bytes64 = (v: string) => Uint8Array.from(atob(v), (c) => c.charCodeAt(0));
export function unpack(bytes: Uint8Array): Unzipped {
  if (bytes.byteLength > 150_000_000) throw Error(t('文件超过 150 MB。', 'File exceeds 150 MB.'));
  let total = 0,
    count = 0;
  return unzipSync(bytes, {
    filter: (f) => {
      total += f.originalSize;
      if (++count > 20000 || total > 400_000_000 || f.originalSize > 100_000_000)
        throw Error(t('文件解压后过大。', 'The expanded package is too large.'));
      if (f.name.startsWith('/') || f.name.includes('\\') || f.name.split('/').includes('..'))
        throw Error(t('文件部件路径无效。', 'Invalid package part path.'));
      return !f.name.endsWith('/');
    },
  });
}
export function resolvePart(from: string, target: string) {
  if (!target || /[?#\\]/.test(target) || /^\w+:/.test(target))
    throw Error(t('不支持此文件引用。', 'Unsupported package relationship.'));
  const parts = target.startsWith('/') ? [] : from.split('/').slice(0, -1);
  for (const s of target.split('/')) {
    if (s === '..') {
      if (!parts.length) throw Error(t('部件路径无效。', 'Invalid part path.'));
      parts.pop();
    } else if (s && s !== '.') parts.push(s);
  }
  return parts.join('/');
}
export function relPath(p: string) {
  const i = p.lastIndexOf('/') + 1;
  return p.slice(0, i) + '_rels/' + p.slice(i) + '.rels';
}
export function relative(from: string, to: string) {
  const a = from.split('/').slice(0, -1),
    b = to.split('/');
  while (a.length && b.length && a[0] === b[0]) {
    a.shift();
    b.shift();
  }
  return [...a.map(() => '..'), ...b].join('/');
}
export class OfficePackage {
  private docs = new Map<string, Document>();
  private dirty = new Set<string>();
  constructor(public files: Unzipped) {}
  doc(path: string) {
    let d = this.docs.get(path);
    if (!d) {
      const b = this.files[path];
      if (!b) throw Error(t('找不到文件部件：', 'Missing package part: ') + path);
      const s = strFromU8(b);
      if (/<!DOCTYPE|<!ENTITY/i.test(s))
        throw Error(t('不支持含实体声明的文档。', 'XML entity declarations are unsupported.'));
      d = new DOMParser().parseFromString(s, 'application/xml');
      if (d.getElementsByTagName('parsererror').length)
        throw Error(t('文件 XML 无效：', 'Invalid XML part: ') + path);
      this.docs.set(path, d);
    }
    return d;
  }
  edit(path: string) {
    this.dirty.add(path);
    return this.doc(path);
  }
  saveDoc(path: string, doc: Document) {
    this.docs.set(path, doc);
    this.dirty.add(path);
  }
  has(path: string) {
    return this.docs.has(path) || Object.prototype.hasOwnProperty.call(this.files, path);
  }
  rels(path: string) {
    return this.has(relPath(path))
      ? Array.from(this.doc(relPath(path)).documentElement.children)
      : [];
  }
  target(path: string, id: string | null) {
    const r = this.rels(path).find((e) => e.getAttribute('Id') === id);
    return r && r.getAttribute('TargetMode') !== 'External'
      ? resolvePart(path, r.getAttribute('Target')!)
      : undefined;
  }
  relate(from: string, to: string, type: string) {
    const rp = relPath(from);
    if (!this.has(rp))
      this.saveDoc(rp, document.implementation.createDocument(REL, 'Relationships'));
    const d = this.edit(rp),
      rs = Array.from(d.documentElement.children),
      target = relative(from, to);
    const found = rs.find(
      (r) => r.getAttribute('Target') === target && r.getAttribute('Type') === type,
    );
    if (found) return found.getAttribute('Id')!;
    const ids = new Set(rs.map((r) => r.getAttribute('Id')));
    let n = 1;
    while (ids.has('rId' + n)) n++;
    const id = 'rId' + n;
    d.documentElement.append(
      element(d, REL, 'Relationship', { Id: id, Type: type, Target: target }),
    );
    return id;
  }
  paths() {
    return [...new Set([...Object.keys(this.files), ...this.docs.keys()])];
  }
  unique(dir: string, stem: string, ext: string) {
    let n = 1;
    while (this.has(dir + '/' + stem + n + ext)) n++;
    return dir + '/' + stem + n + ext;
  }
  contentType(path: string) {
    const entries = Array.from(this.doc('[Content_Types].xml').documentElement.children);
    return (
      entries.find((x) => x.getAttribute('PartName') === '/' + path) ||
      entries.find((x) => x.getAttribute('Extension') === path.split('.').pop())
    )?.getAttribute('ContentType');
  }
  type(path: string, type: string) {
    if (this.contentType(path) === type) return;
    const d = this.edit('[Content_Types].xml');
    d.documentElement.append(
      element(d, CT, 'Override', { PartName: '/' + path, ContentType: type }),
    );
  }
  pack() {
    for (const path of this.dirty) this.files[path] = strToU8(xml(this.doc(path)));
    return zipSync(this.files, { level: 4 });
  }
}
