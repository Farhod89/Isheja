/* ====================================================================
   Фотоҳисобот слайд генератори (PptxGenJS, браузерда, серверсиз)
   3 та шаблон: 'green', 'blue', 'mono'
   ==================================================================== */

function buildSlide(opts) {
  // opts = { template, orgTitle, schoolName, computerCount, doskaCount, photos: [dataUrl,...], note }
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'A4W', width: 10, height: 5.63 });
  pptx.layout = 'A4W';

  const slide = pptx.addSlide();
  const T = opts.template || 'green';
  const photos = (opts.photos || []).slice(0, 4);

  const palettes = {
    green: { bg: 'FFFFFF', bar: '14532D', accent: '4ADE80', text: '14532D', sub: '4B6358' },
    blue:  { bg: 'F3F7FB', bar: '1E3A8A', accent: '60A5FA', text: '1E3A8A', sub: '45587A' },
    mono:  { bg: 'FFFFFF', bar: '1F2937', accent: '9CA3AF', text: '111827', sub: '4B5563' },
  };
  const c = palettes[T] || palettes.green;

  slide.background = { color: c.bg };

  // header bar
  slide.addShape('rect', { x: 0, y: 0, w: 10, h: 0.9, fill: { color: c.bar } });
  slide.addText(opts.orgTitle || 'Ёзёвон тумани Мактабгача ва мактаб таълими бўлими', {
    x: 0.3, y: 0.08, w: 9.4, h: 0.4, fontSize: 14, bold: true, color: 'FFFFFF', fontFace: 'Calibri'
  });
  slide.addText('Компьютер техникалари ва электрон доскалар бўйича фотоҳисобот', {
    x: 0.3, y: 0.46, w: 9.4, h: 0.34, fontSize: 11, color: 'E8F3EC', fontFace: 'Calibri'
  });

  // school name + stats row
  slide.addText(opts.schoolName || '[Мактаб номи]', {
    x: 0.3, y: 1.05, w: 6.0, h: 0.4, fontSize: 18, bold: true, color: c.text, fontFace: 'Calibri'
  });

  const statBoxW = 1.6;
  slide.addShape('roundRect', { x: 6.5, y: 1.0, w: statBoxW, h: 0.6, fill: { color: c.bar }, rectRadius: 0.08 });
  slide.addText(`${opts.computerCount ?? 8}\nкомпьютер`, {
    x: 6.5, y: 1.0, w: statBoxW, h: 0.6, fontSize: 11, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle'
  });
  slide.addShape('roundRect', { x: 8.2, y: 1.0, w: statBoxW, h: 0.6, fill: { color: c.accent }, rectRadius: 0.08 });
  slide.addText(`${opts.doskaCount ?? 1}\nэлектрон доска`, {
    x: 8.2, y: 1.0, w: statBoxW, h: 0.6, fontSize: 11, bold: true, color: T === 'mono' ? '111827' : '14532D', align: 'center', valign: 'middle'
  });

  // photo grid
  const gridY = 1.8, gridH = 3.5, gap = 0.12;
  const n = Math.max(photos.length, 1);
  const cols = n <= 1 ? 1 : 2;
  const rows = n <= 2 ? 1 : 2;
  const cellW = (9.4 - gap * (cols - 1)) / cols;
  const cellH = (gridH - gap * (rows - 1)) / rows;

  if (photos.length === 0) {
    slide.addShape('rect', { x: 0.3, y: gridY, w: 9.4, h: gridH, fill: { color: 'EFF3F0' }, line: { color: 'CBD5C9', dashType: 'dash' } });
    slide.addText('Расм юкланмаган — форма орқали фото қўшинг', {
      x: 0.3, y: gridY, w: 9.4, h: gridH, align: 'center', valign: 'middle', color: '8A9A91', fontSize: 14
    });
  } else {
    photos.forEach((dataUrl, i) => {
      const col = i % cols, row = Math.floor(i / cols);
      const x = 0.3 + col * (cellW + gap);
      const y = gridY + row * (cellH + gap);
      slide.addImage({ data: dataUrl, x, y, w: cellW, h: cellH, sizing: { type: 'cover', w: cellW, h: cellH } });
      slide.addShape('rect', { x, y, w: cellW, h: cellH, fill: { type: 'none' }, line: { color: c.bar, width: 1 } });
    });
  }

  if (opts.note) {
    slide.addText(opts.note, { x: 0.3, y: 5.35, w: 9.4, h: 0.25, fontSize: 9, italic: true, color: c.sub });
  }

  return pptx;
}

async function generateAndDownload(opts) {
  const pptx = buildSlide(opts);
  const safeName = (opts.schoolName || 'maktab').replace(/[^\p{L}\p{N}]+/gu, '_');
  await pptx.writeFile({ fileName: `fotohisobot_${safeName}.pptx` });
}
