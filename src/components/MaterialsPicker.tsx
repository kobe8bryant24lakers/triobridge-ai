import { englishCount } from './presentation-runtime';
import { t } from '../i18n';
import { useRef, useState } from 'react';
import { FolderOpen, FilePlus2, X } from 'lucide-react';
import { readMaterials } from '../presentation/materials';
import type { Material, MaterialContext, MaterialIssue } from '../presentation/types';

interface Props {
  context: MaterialContext;
  onChange: (value: MaterialContext) => void;
  busy: boolean;
  run: (
    label: string,
    work: (signal: AbortSignal) => Promise<void>,
    cancellable?: boolean,
  ) => Promise<void>;
}
export function MaterialsPicker({ context, onChange, busy, run }: Props) {
  const [materials, setMaterials] = useState<Material[]>([]),
    [issues, setIssues] = useState<MaterialIssue[]>([]),
    [progress, setProgress] = useState('');
  const folder = useRef<HTMLInputElement>(null),
    files = useRef<HTMLInputElement>(null);
  const [open, setOpen] = useState(true);
  async function choose(list: FileList | null) {
    if (!list?.length) return;
    const chosen = Array.from(list);
    await run(
      t('读取本地素材', 'Reading local files'),
      async (signal) => {
        try {
          const result = await readMaterials(chosen, signal, (name, i, total) =>
            setProgress(`${i} / ${total} · ${name}`),
          );
          setMaterials(result.materials);
          setIssues(result.issues);
          onChange({ materials: result.materials, includeImages: false });
          setOpen(true);
        } finally {
          setProgress('');
        }
      },
      true,
    );
  }
  const selected = new Set(context.materials.map((m) => m.id));
  return (
    <section className="ppt-materials" aria-label={t('素材上下文', 'Source files')}>
      <div className="section-heading">
        <button
          type="button"
          className="text-button"
          aria-expanded={open}
          onClick={() => setOpen(!open)}
        >
          <FolderOpen size={17} />
          {t('素材上下文 · ', 'Source files · ')}
          {context.materials.length}
          {t(' 份已选', ' selected')}
        </button>
        {materials.length > 0 && (
          <button
            type="button"
            className="icon-button"
            aria-label={t('清空素材', 'Clear sources')}
            disabled={busy}
            onClick={() => {
              setMaterials([]);
              setIssues([]);
              onChange({ materials: [], includeImages: false });
            }}
          >
            <X size={16} />
          </button>
        )}
      </div>
      {open && (
        <>
          <p className="field-help">
            {t(
              '选择本地目录，或挑选文件。先在本机解析，点击生成时才将勾选内容发送给当前模型。再次选择会替换本次素材。',
              'Choose a folder or files. Files are read locally. Generate sends checked files to the model. Choosing again replaces this selection.',
            )}
          </p>
          <div className="button-row">
            <button
              type="button"
              className="secondary"
              disabled={busy}
              onClick={() => folder.current?.click()}
            >
              <FolderOpen size={16} />
              {t('选择素材目录', 'Choose folder')}
            </button>
            <button
              type="button"
              className="secondary"
              disabled={busy}
              onClick={() => files.current?.click()}
            >
              <FilePlus2 size={16} />
              {t('选择文件', 'Choose files')}
            </button>
          </div>
          <input
            ref={folder}
            className="sr-only"
            type="file"
            multiple
            aria-label={t('选择素材目录文件', 'Files in source folder')}
            tabIndex={-1}
            {...{ webkitdirectory: '', directory: '' }}
            onChange={(e) => {
              void choose(e.target.files);
              e.target.value = '';
            }}
          />
          <input
            ref={files}
            className="sr-only"
            type="file"
            multiple
            accept=".txt,.md,.csv,.tsv,.json,.docx,.xlsx,.pptx,.pdf,.png,.jpg,.jpeg,.webp"
            aria-label={t('选择素材文件', 'Select source files')}
            tabIndex={-1}
            onChange={(e) => {
              void choose(e.target.files);
              e.target.value = '';
            }}
          />
          <p className="field-help">
            {t(
              'Word、Excel、PPT、PDF、文本与图片。最多 40 个文件 / 80 MB，单文件 20 MB。不自动读取目录外的文件。',
              'Word, Excel, PowerPoint, PDF, text and images. Up to 40 files, 20 MB each and 80 MB total. Only selected files are read.',
            )}
          </p>
          {progress && (
            <p role="status" className="field-help">
              {progress}
            </p>
          )}
          {materials.length > 0 && (
            <div className="material-list">
              {materials.map((m) => (
                <div className="material-row" key={m.id}>
                  <label className="material-choice">
                    <input
                      type="checkbox"
                      checked={selected.has(m.id)}
                      disabled={busy}
                      onChange={(e) =>
                        onChange({
                          ...context,
                          materials: e.target.checked
                            ? materials.filter((item) => selected.has(item.id) || item.id === m.id)
                            : context.materials.filter((item) => item.id !== m.id),
                        })
                      }
                    />
                    <span>
                      {m.name}
                      <small>
                        {m.image
                          ? t('图片', 'Image')
                          : t(
                              `${m.text.length.toLocaleString()} 字符`,
                              `${m.text.length.toLocaleString()} characters`,
                            )}
                        {m.tables.length > 0
                          ? t(
                              ` · ${m.tables.length} 个表格`,
                              ` · ${englishCount(m.tables.length, 'table')}`,
                            )
                          : ''}
                      </small>
                    </span>
                  </label>
                  <details>
                    <summary>{t('查看内容', 'View content')}</summary>
                    {m.image ? (
                      <img className="material-image" src={m.image.modelData} alt={m.name} />
                    ) : (
                      <pre className="material-text">
                        {m.text.slice(0, 2400)}
                        {m.text.length > 2400
                          ? t(
                              '\n… 预览仅显示前 2,400 字符，发送时包含完整已提取内容。',
                              '\n… The preview shows the first 2,400 characters. The full extracted content will be sent.',
                            )
                          : ''}
                      </pre>
                    )}
                    {m.warning && <p className="field-help">{m.warning}</p>}
                  </details>
                </div>
              ))}
            </div>
          )}
          {materials.some((m) => m.image) && (
            <label className="checkbox-label">
              <input
                type="checkbox"
                checked={context.includeImages}
                disabled={busy}
                onChange={(e) => onChange({ ...context, includeImages: e.target.checked })}
              />
              {t(
                '将图片预览也发给模型理解（最多 4 张，需视觉模型）',
                'Send image previews to the model (up to 4, requires a vision model)',
              )}
            </label>
          )}
          {materials.some((m) => m.image) && !context.includeImages && (
            <p className="field-help">
              {t(
                '图片只按名称供模型选用，原图经本机压缩后嵌入 PPT；不会读取图片中的事实。',
                'Images are identified by filename for layout and compressed locally for the deck. Their content is not analyzed.',
              )}
            </p>
          )}
          {issues.length > 0 && (
            <details className="material-issues" open={!materials.length}>
              <summary>
                {t(`${issues.length} 项资料读取提示`, englishCount(issues.length, 'file notice'))}
              </summary>
              <ul>
                {issues.map((issue, i) => (
                  <li key={i}>
                    <strong>{issue.name}</strong>
                    {t('：', ': ')}
                    {issue.message}
                  </li>
                ))}
              </ul>
            </details>
          )}
        </>
      )}
    </section>
  );
}
