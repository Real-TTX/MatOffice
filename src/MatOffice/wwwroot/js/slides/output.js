/* MatOffice Slides – what leaves the editor in the browser: a slide as a PNG (the thumbnail of the gallery, "Download as PNG"), drawn by putting the
   DOM of render.js into an SVG foreignObject and that into a canvas; and printing: full-page slides (landscape), notes pages or handouts with
   three slides per page (portrait), each page built in a print container that only shows while printing. */
import { t, el } from '../editor/util.js';
import { renderSlide, fitText, RENDER_CSS } from './render.js';

/** A slide as a PNG blob, width px wide. */
export async function slidePng(content, slide, number, width) {
  const { w, h } = content.size;
  const height = Math.round((h * width) / w);
  const node = renderSlide(content, slide, { mode: 'view', number });
  const host = el('div', { class: 'mo-slides-offscreen', style: { position: 'fixed', left: '-30000px', top: '0', width: w + 'px', height: h + 'px', overflow: 'hidden', pointerEvents: 'none' } });
  host.append(node);
  document.body.append(host);
  try {
    fitText(node);
    if (document.fonts?.ready) await document.fonts.ready;
    await Promise.all([...node.querySelectorAll('img')].map(i => i.decode().catch(() => {})));
    const xhtml = new XMLSerializer().serializeToString(node);
    const css = RENDER_CSS.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const svgText = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${w} ${h}"><foreignObject x="0" y="0" width="${w}" height="${h}"><div xmlns="http://www.w3.org/1999/xhtml" style="width:${w}px;height:${h}px"><style>${css}</style>${xhtml}</div></foreignObject></svg>`;
    const url = URL.createObjectURL(new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' }));
    try {
      const img = new Image();
      img.decoding = 'sync';
      img.src = url;
      await img.decode();
      const canvas = el('canvas', { width, height });
      const g = canvas.getContext('2d');
      g.fillStyle = '#ffffff';
      g.fillRect(0, 0, width, height);
      g.drawImage(img, 0, 0, width, height);
      return await new Promise((resolve, reject) => canvas.toBlob(b => (b ? resolve(b) : reject(new Error('png'))), 'image/png'));
    } finally {
      URL.revokeObjectURL(url);
    }
  } finally {
    host.remove();
  }
}

/** Prints the presentation: kind 'slides' (one per landscape page), 'notes' (slide and notes), 'handouts' (three per page with lines). */
export async function printDeck(content, kind = 'slides', slidesToPrint = null) {
  document.querySelector('.mo-slides-print')?.remove();
  document.getElementById('mo-slides-page')?.remove();
  const { w, h } = content.size;
  const list = slidesToPrint || content.slides.map((s, i) => ({ slide: s, number: i + 1 })).filter(x => !x.slide.hidden);
  const host = el('div', { class: 'mo-slides-print mo-slides-print--' + kind });
  const portrait = kind !== 'slides';
  const page = el('style', { id: 'mo-slides-page', text: `@page { size: ${portrait ? 'portrait' : 'landscape'}; margin: 0.45in; }` });
  // the printable area of A4/Letter with these margins: a slide 9.6 in wide (landscape), 6.6 in (portrait)
  const slideWidth = kind === 'slides' ? Math.min(9.6 * 96, (6.8 * 96 * w) / h) : kind === 'notes' ? 6.4 * 96 : 3.3 * 96;
  const frame = (slide, number) => {
    const scale = slideWidth / w;
    const node = renderSlide(content, slide, { mode: 'view', number });
    node.style.transform = `scale(${scale})`;
    node.style.transformOrigin = '0 0';
    return el('div', { class: 'mo-slides-print__slide', style: { width: slideWidth + 'px', height: (h * scale) + 'px' } }, node);
  };
  if (kind === 'handouts') {
    for (let i = 0; i < list.length; i += 3) {
      const sheet = el('section', { class: 'mo-slides-print__page' });
      for (const x of list.slice(i, i + 3)) {
        const lines = el('div', { class: 'mo-slides-print__lines' });
        for (let k = 0; k < 7; k++) lines.append(el('span'));
        sheet.append(el('div', { class: 'mo-slides-print__handout' }, frame(x.slide, x.number), lines));
      }
      host.append(sheet);
    }
  } else {
    for (const x of list) {
      const sheet = el('section', { class: 'mo-slides-print__page' }, frame(x.slide, x.number));
      if (kind === 'notes') sheet.append(el('div', { class: 'mo-slides-print__notes', text: x.slide.notes || '' }), el('div', { class: 'mo-slides-print__num', text: String(x.number) }));
      host.append(sheet);
    }
  }
  document.head.append(page);
  document.body.append(host);
  document.body.classList.add('mo-slides-printing');
  host.querySelectorAll('.mo-sr').forEach(fitText);
  await Promise.all([...host.querySelectorAll('img')].map(i => i.decode().catch(() => {})));
  const cleanup = () => { host.remove(); page.remove(); document.body.classList.remove('mo-slides-printing'); window.removeEventListener('afterprint', cleanup); };
  window.addEventListener('afterprint', cleanup);
  window.print();
  setTimeout(() => { if (!matchMedia('print').matches) cleanup(); }, 1500);
}

/** The choice of the print layout (like the print settings of the office suites); resolves with the kind or null. */
export async function askPrint(dialog) {
  const options = [['slides', t('Slides.Print.Slides')], ['notes', t('Slides.Print.Notes')], ['handouts', t('Slides.Print.Handouts')]];
  const body = el('div', { class: 'mo-slides-printchoice' }, ...options.map(([value, label], i) => el('label', { class: 'mat-check' },
    el('input', { type: 'radio', name: 'kind', value, checked: i === 0 ? true : null }), el('span', { text: label }))));
  let chosen = 'slides';
  body.addEventListener('change', e => { chosen = e.target.value; });
  const r = await dialog({ title: t('Slides.Print.Title'), icon: 'print', body, buttons: [{ label: t('Slides.Print.Button'), value: 'print', kind: 'primary', icon: 'print' }, { label: t('Editor.Cancel'), value: 'cancel' }] });
  return r === 'print' ? chosen : null;
}
