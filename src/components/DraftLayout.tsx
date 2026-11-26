import { useState } from 'react';
import { PanelsTopLeft } from 'lucide-react';
import { t } from '../i18n';
import { arrangementName, suggestedArrangement } from '../presentation/layout-optimization';
import { SlidePreview } from './PresentationPreview';
import type { DeckContent, MaterialContext, PresentationTheme } from '../presentation/types';
export function DraftLayout({
  deck,
  busy,
  onApply,
  context,
  theme,
}: {
  context: MaterialContext;
  theme: PresentationTheme;
  deck: DeckContent;
  busy: boolean;
  onApply: (deck: DeckContent) => void;
}) {
  const [review, setReview] = useState<DeckContent>(),
    [chosen, setChosen] = useState<number[]>([]);
  const available = deck.slides.flatMap((s, i) => {
    const layout = suggestedArrangement(s);
    return layout && layout !== s.arrangement ? [{ i, layout }] : [];
  });
  const [previewPage, setPreviewPage] = useState(0);
  const previewSlide = available.find((v) => v.i === previewPage) || available[0];
  const visible = review === deck;
  return (
    <section className="draft-layout">
      <button
        type="button"
        className="secondary"
        disabled={busy}
        aria-expanded={visible}
        onClick={() => {
          setReview(visible ? undefined : deck);
          setChosen(available.map((x) => x.i));
        }}
      >
        <PanelsTopLeft size={16} />
        {t('优化布局', 'Optimize layout')}
      </button>
      {visible && (
        <div className="theme-sheet">
          <p className="field-help">
            {available.length
              ? t(
                  '勾选需要换版式的页，文字和数据沿用原稿。',
                  'Choose slides to reflow using their existing text and data.',
                )
              : t('当前版式没有可自动重排的页。', 'There are no eligible slides to reflow.')}
          </p>
          {available.map(({ i, layout }) => (
            <label className="checkbox-label" key={i}>
              <input
                type="checkbox"
                disabled={busy}
                checked={chosen.includes(i)}
                onChange={(e) =>
                  setChosen(e.target.checked ? [...chosen, i] : chosen.filter((v) => v !== i))
                }
              />
              <span>
                {i + 1}. {deck.slides[i].title} · {arrangementName(layout)}
              </span>
            </label>
          ))}
          {previewSlide && (
            <>
              <label>
                {t('重排效果预览', 'Reflow preview')}
                <select
                  value={previewSlide.i}
                  disabled={busy}
                  onChange={(e) => setPreviewPage(Number(e.target.value))}
                >
                  {available.map((v) => (
                    <option value={v.i} key={v.i}>
                      {v.i + 1}. {deck.slides[v.i].title}
                    </option>
                  ))}
                </select>
              </label>
              <SlidePreview
                page={{ ...deck.slides[previewSlide.i], arrangement: previewSlide.layout }}
                context={context}
                theme={theme}
              />
            </>
          )}
          {available.length > 0 && (
            <button
              type="button"
              className="primary"
              disabled={busy || !chosen.length}
              onClick={() => {
                onApply({
                  ...deck,
                  slides: deck.slides.map((s, i) =>
                    chosen.includes(i) ? { ...s, arrangement: suggestedArrangement(s) } : s,
                  ),
                });
                setReview(undefined);
              }}
            >
              {t(
                '应用所选 ' + chosen.length + ' 页布局',
                'Apply layouts to ' + chosen.length + ' slides',
              )}
            </button>
          )}
        </div>
      )}
    </section>
  );
}
