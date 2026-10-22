/* Заполняет template.docx данными, полученными от Gemini. Работает в браузере и в Node. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DocBuilder = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const LANG = { kk: 'kk-KZ', ru: 'ru-RU' };
  const TYPES = {
    combined: { kk: 'Аралас', ru: 'Комбинированный' },
    practice: { kk: 'Практикалық', ru: 'Практическое занятие' }
  };
  const PPR = b => `<w:pPr><w:pStyle w:val="ac"/><w:spacing w:before="${b || 0}" w:beforeAutospacing="0" w:after="0" w:afterAutospacing="0"/>`;
  const clean = s => String(s == null ? '' : s).replace(/^\s*(?:[•\-–]\s+|\d+[.)]\s+)/, '').trim();

  function runs(text, lang, bold, ital, sz) {
    return String(text).split(/(\*\*[^*]+\*\*|`[^`]+`|\*[^*\s][^*]*\*)/).filter(Boolean).map(part => {
      let b = bold, i = ital, code = false, t = part;
      if (/^\*\*.+\*\*$/.test(part)) { b = true; t = part.slice(2, -2); }
      else if (/^`.+`$/.test(part)) { code = true; t = part.slice(1, -1); }
      else if (/^\*.+\*$/.test(part)) { i = true; t = part.slice(1, -1); }
      const rpr = '<w:rPr>' + (code ? '<w:rStyle w:val="HTML"/>' : '') + (b ? '<w:b/><w:bCs/>' : '') + (i ? '<w:i/><w:iCs/>' : '') +
        (sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : '') + `<w:lang w:val="${lang}"/></w:rPr>`;
      return `<w:r>${rpr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
    }).join('');
  }
  const para = (inner, extra, before) => `<w:p>${PPR(before)}${extra || ''}</w:pPr>${inner}</w:p>`;
  const item = (mark, text, lang) => para(`<w:r><w:t>${mark}</w:t></w:r><w:r><w:tab/></w:r>` + runs(text, lang),
    '<w:ind w:left="340" w:hanging="340"/>');

  function table(b, lang) {
    const heads = (b.headers || []).map(String), rows = (b.rows || []).map(r => (r.cells || r || []).map(String));
    const n = Math.max(heads.length, ...rows.map(r => r.length), 1), total = 3885;
    const w = n === 3 ? [1200, 1050, 1635] : Array(n).fill(Math.floor(total / n));
    const bd = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(k => `<w:${k} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('');
    const cell = (t, i, head) => `<w:tc><w:tcPr><w:tcW w:w="${w[i]}" w:type="dxa"/>${head ? '<w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/>' : ''}</w:tcPr>` +
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>${runs(t || '', lang, head, false, 18)}</w:p></w:tc>`;
    const row = (cs, head) => '<w:tr>' + Array.from({ length: n }, (_, i) => cell(cs[i], i, head)).join('') + '</w:tr>';
    return `<w:tbl><w:tblPr><w:tblW w:w="${w.reduce((a, c) => a + c, 0)}" w:type="dxa"/><w:tblBorders>${bd}</w:tblBorders><w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${w.map(x => `<w:gridCol w:w="${x}"/>`).join('')}</w:tblGrid>` +
      (heads.length ? row(heads, true) : '') + rows.map(r => row(r, false)).join('') + '</w:tbl>' + para('', '');
  }

  function blocks(list, lang) {
    let num = 0, out = '';
    (list || []).forEach(b => {
      if (b.kind !== 'numbered') num = 0;
      const t = clean(b.text);
      if (b.kind === 'heading') out += para(runs(String(b.text || '').trim(), lang, true), '', 80);
      else if (b.kind === 'bullet') out += item('•', t, lang);
      else if (b.kind === 'numbered') out += item(++num + '.', t, lang);
      else if (b.kind === 'quote') out += para(runs(t, lang, false, true));
      else if (b.kind === 'table') out += table(b, lang);
      else out += para(runs(String(b.text || '').trim(), lang));
    });
    return out || para('', '');
  }
  const lines = (arr, mark, lang, prefix) => (arr || []).map(s => item(mark, (prefix || '') + clean(s), lang)).join('');
  const dmy = iso => (iso || '').split('-').reverse().join('.');

  async function build(JSZip, templateBuf, form, d) {
    const lang = LANG[form.lang] || 'ru-RU';
    const zip = await JSZip.loadAsync(templateBuf);
    let xml = await zip.file('word/document.xml').async('string');
    const text = { MODULE: form.module, TOPIC: form.topic, FIO: form.fio, DATE: dmy(form.date), GROUP: form.group,
      LESSON_TYPE: (TYPES[form.lessonType] || TYPES.combined)[form.lang] || TYPES.combined.ru,
      START_ASSESS: d.start && d.start.assessment, START_RES: d.start && d.start.resources,
      MIDDLE_ASSESS: d.middle && d.middle.assessment, MIDDLE_RES: d.middle && d.middle.resources,
      END_ASSESS: d.end && d.end.assessment, END_RES: d.end && d.end.resources };
    const blk = { GOALS: lines(d.goals, '•', lang), RESULTS: lines(d.results, '•', lang), CHECKLIST: lines(d.checklist, '•', lang, '[ ] '),
      RESOURCES: para(runs(d.resources || '', lang)), START: blocks(d.start && d.start.blocks, lang),
      MIDDLE: blocks(d.middle && d.middle.blocks, lang), END: blocks(d.end && d.end.blocks, lang) };
    Object.keys(blk).forEach(k => { xml = xml.split(`<w:p><w:r><w:t>{{${k}}}</w:t></w:r></w:p>`).join(blk[k]); });
    Object.keys(text).forEach(k => { xml = xml.split(`{{${k}}}`).join(esc(text[k] || '')); });
    if (/\{\{\w+\}\}/.test(xml)) throw new Error('Шаблон заполнен не полностью');
    zip.file('word/document.xml', xml);
    return zip.generateAsync({ type: typeof Blob !== 'undefined' ? 'blob' : 'nodebuffer', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }
  return { build, dmy, TYPES };
});
