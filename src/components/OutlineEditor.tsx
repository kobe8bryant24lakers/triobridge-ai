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
      <p className="field-help">每页一个标题，1–5 条要点。每行一条要点，可调整页面顺序。</p>
      {slides.map((slide, index) => (
        <fieldset className="slide-editor" key={index} disabled={disabled}>
          <legend>第 {index + 1} 页</legend>
          <div className="slide-edit-actions">
            <button
              className="icon-button"
              aria-label={`上移第 ${index + 1} 页`}
              disabled={disabled || index === 0}
              onClick={() => move(index, -1)}
            >
              <ArrowUp size={16} />
            </button>
            <button
              className="icon-button"
              aria-label={`下移第 ${index + 1} 页`}
              disabled={disabled || index === slides.length - 1}
              onClick={() => move(index, 1)}
            >
              <ArrowDown size={16} />
            </button>
            <button
              className="icon-button"
              aria-label={`删除第 ${index + 1} 页`}
              disabled={disabled || slides.length === 1}
              onClick={() => onChange(slides.filter((_, i) => i !== index))}
            >
              <Trash2 size={16} />
            </button>
          </div>
          <label>
            第 {index + 1} 页标题
            <input
              value={slide.title}
              aria-invalid={!slide.title.trim() || slide.title.length > 60}
              onChange={(e) => update(index, { title: e.target.value })}
            />
          </label>
          <span className="field-help">{slide.title.length} / 60 字符</span>
          <label>
            第 {index + 1} 页要点
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
          <span className="field-help">{slide.bullets.length} / 5 条，每条最多 160 字符</span>
        </fieldset>
      ))}
      <button
        className="secondary"
        disabled={disabled || slides.length >= 12}
        onClick={() => onChange([...slides, { title: '新页面', bullets: ['待补充要点'] }])}
      >
        <Plus size={16} />
        添加一页（{slides.length} / 12）
      </button>
    </div>
  );
}
