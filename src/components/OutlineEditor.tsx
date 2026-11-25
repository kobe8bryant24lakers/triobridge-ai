import { t } from '../i18n';
import { ArrowDown, ArrowUp, Plus, Trash2 } from 'lucide-react';
import type { SlideContent } from '../core/types';

interface Props {
  slides: SlideContent[];
  disabled: boolean;
  onChange: (slides: SlideContent[]) => void;
}

export function OutlineEditor({ slides, disabled, onChange }: Props) {
  function update(index: number, change: Partial<SlideContent>) {
    onChange(slides.map((slide, i) => (i === index ? { ...slide, ...change } : slide)));
  }
  function move(index: number, offset: number) {
    const next = [...slides];
    [next[index], next[index + offset]] = [next[index + offset], next[index]];
    onChange(next);
  }
  return (
    <div className="outline-editor">
      <p className="field-help">
        {t(
          '每页一个标题，1–5 条要点。每行一条要点，可调整页面顺序。',
          'Each slide has a title and 1–5 bullets. Enter one bullet per line. You can reorder slides.',
        )}
      </p>
      {slides.map((slide, index) => (
        <fieldset className="slide-editor" key={index} disabled={disabled}>
          <legend>{t(`第 ${index + 1} 页`, `Slide ${index + 1}`)}</legend>
          <div className="slide-edit-actions">
            <button
              className="icon-button"
              aria-label={t(`上移第 ${index + 1} 页`, `Move slide ${index + 1} up`)}
              disabled={disabled || index === 0}
              onClick={() => move(index, -1)}
            >
              <ArrowUp size={16} />
            </button>
            <button
              className="icon-button"
              aria-label={t(`下移第 ${index + 1} 页`, `Move slide ${index + 1} down`)}
              disabled={disabled || index === slides.length - 1}
              onClick={() => move(index, 1)}
            >
              <ArrowDown size={16} />
            </button>
            <button
              className="icon-button"
              aria-label={t(`删除第 ${index + 1} 页`, `Delete slide ${index + 1}`)}
              disabled={disabled || slides.length === 1}
              onClick={() => onChange(slides.filter((_, i) => i !== index))}
            >
              <Trash2 size={16} />
            </button>
          </div>
          <label>
            {t(`第 ${index + 1} 页标题`, `Slide ${index + 1} title`)}
            <input
              value={slide.title}
              aria-invalid={!slide.title.trim() || slide.title.length > 60}
              onChange={(e) => update(index, { title: e.target.value })}
            />
          </label>
          <span className="field-help">
            {slide.title.length}
            {t(' / 60 字符', ' / 60 characters')}
          </span>
          <label>
            {t(`第 ${index + 1} 页要点`, `Slide ${index + 1} bullets`)}
            <textarea
              rows={Math.max(3, Math.min(6, slide.bullets.length))}
              value={slide.bullets.join('\n')}
              onChange={(e) => update(index, { bullets: e.target.value.split('\n') })}
              aria-invalid={
                slide.bullets.length > 5 ||
                slide.bullets.some((bullet) => !bullet.trim() || bullet.length > 160)
              }
            />
          </label>
          <span className="field-help">
            {slide.bullets.length}
            {t(' / 5 条，每条最多 160 字符', ' / 5 bullets, up to 160 characters each')}
          </span>
        </fieldset>
      ))}
      <button
        className="secondary"
        disabled={disabled || slides.length >= 12}
        onClick={() =>
          onChange([
            ...slides,
            { title: t('新页面', 'New slide'), bullets: [t('待补充要点', 'Add a bullet')] },
          ])
        }
      >
        <Plus size={16} />
        {t(`添加一页（${slides.length} / 12）`, `Add slide (${slides.length} / 12)`)}
      </button>
    </div>
  );
}
