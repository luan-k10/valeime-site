// valei-me · recebe as sugestões do site (valeime.netlify.app) e organiza a planilha.
// Abas: 📊 Resumo (contadores + rankings sem repetidas), 🎵 Músicas, ✦ Santos, ✝ Posts.
// Menu "valei-me" dentro da planilha: Atualizar resumo / Refazer formatação.

const TZ = 'America/Sao_Paulo';
const ADMIN_KEY = 'TROQUE-POR-UMA-CHAVE-SECRETA';
const SH = { res: '📊 Resumo', mus: '🎵 Músicas', santo: '✦ Santos', post: '✝ Posts' };
const C = {
  ink: '#1E1C20', gold: '#C6921C', rose: '#CE506C', wine: '#9E2A35', cream: '#FAF8F4',
  line: '#E8E1D5', muted: '#9A948C', roseBg: '#F9E3E8', wineBg: '#F6E1E3', doneBg: '#EFEBE4',
  white: '#FFFFFF',
};
const FONT = 'Inter';
const MUS_HEAD = ['Feito', 'Data', 'Música', 'Artista', 'Pedidos', 'Link', 'Trecho',
  'Por que essa música', 'Quem sugeriu', 'Instagram', 'chave'];
const SANTO_HEAD = ['Feito', 'Data', 'Santo', 'Pedidos', 'Por que esse santo', 'Quem sugeriu',
  'Instagram', 'chave'];
const POST_HEAD = ['Feito', 'Data', 'Ideia de post', 'Quem sugeriu', 'Instagram'];

// Abas com contagem de pedidos repetidos (agrupados por uma chave normalizada).
const RANKED = {
  mus: { sheet: SH.mus, keyCol: 11, titleCol: 3, countCol: 5, subCol: 4 },
  santo: { sheet: SH.santo, keyCol: 8, titleCol: 3, countCol: 4, subCol: 0 },
};

// Resumo: ranking de músicas em B:G e de santos em I:L; chaves ocultas em N e O.
const RANK_START = 11;
const RES = { musDone: 7, santoDone: 12, musKey: 14, santoKey: 15, cols: 15 };

const ss = () => SpreadsheetApp.getActiveSpreadsheet();
const checkboxRule = () => SpreadsheetApp.newDataValidation().requireCheckbox().build();
const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
// "São Miguel", "Sao Miguel" e "Santo Miguel" viram a mesma chave.
const normSanto = s => norm(s).replace(/^(sao|santo|santa|sto|sta|s)\s+/, '');
const keyOf = (kind, s) => (kind === 'santo' ? normSanto(s) : norm(s));
const text = s => ContentService.createTextOutput(s);

// ---------- recebimento ----------

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = (e && e.parameter) || {};
    if (p.site) return json({ ok: true });
    if (!ss().getSheetByName(SH.mus) || !ss().getSheetByName(SH.santo)) setup();
    const clean = v => String(v || '').trim().slice(0, 2000).replace(/^[=+\-@]/, "'$&");
    const date = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');

    if (p.tipo === 'Ideia de post') {
      const sh = ss().getSheetByName(SH.post);
      const r = nextRow(sh);
      sh.getRange(r, 1, 1, 5).setValues([[false, date, clean(p.ideia), clean(p.nome), clean(p.instagram)]]);
      sh.getRange(r, 1).setDataValidation(checkboxRule());
    } else if (p.tipo === 'Santo da semana') {
      const sh = ss().getSheetByName(SH.santo);
      const key = keyOf('santo', p.santo);
      const done = key ? alreadyDone(sh, key, RANKED.santo.keyCol) : false;
      const r = nextRow(sh);
      sh.getRange(r, 1, 1, 8).setValues([[done, date, clean(p.santo), '', clean(p.ideia),
        clean(p.nome), clean(p.instagram), key]]);
      sh.getRange(r, 1).setDataValidation(checkboxRule());
    } else {
      const sh = ss().getSheetByName(SH.mus);
      const key = keyOf('mus', p.musica);
      const done = key ? alreadyDone(sh, key, RANKED.mus.keyCol) : false;
      const r = nextRow(sh);
      sh.getRange(r, 1, 1, 11).setValues([[done, date, clean(p.musica), clean(p.artista), '',
        clean(p.link), clean(p.trecho), clean(p.ideia), clean(p.nome), clean(p.instagram), key]]);
      sh.getRange(r, 1).setDataValidation(checkboxRule());
    }
    refreshRanking();
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.k === ADMIN_KEY) {
    if (p.a === 'setup') return text(setup());
    if (p.a === 'limpar') return text(limparTestes());
    if (p.a === 'status') return json(status());
  }
  return text('valei-me: recebendo sugestoes. Pode fechar esta aba.');
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function status() {
  const rows = name => {
    const sh = ss().getSheetByName(name);
    return sh ? Math.max(sh.getLastRow() - 1, 0) : null;
  };
  const res = ss().getSheetByName(SH.res);
  return {
    abas: ss().getSheets().map(s => s.getName()),
    musicas: rows(SH.mus), santos: rows(SH.santo), posts: rows(SH.post),
    contadores: res ? res.getRange('B6:L6').getValues()[0] : null,
  };
}

function nextRow(sh) {
  const r = sh.getLastRow() + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  return r;
}

function alreadyDone(sh, key, keyCol) {
  const last = sh.getLastRow();
  if (last < 2) return false;
  const a = sh.getRange(2, 1, last - 1, 1).getValues();
  const k = sh.getRange(2, keyCol, last - 1, 1).getValues();
  return k.some((row, i) => row[0] === key && a[i][0] === true);
}

// ---------- planilha: gatilhos e menu ----------

function onOpen() {
  SpreadsheetApp.getUi().createMenu('valei-me')
    .addItem('Atualizar resumo', 'refreshRanking')
    .addItem('Refazer formatação', 'setup')
    .addToUi();
  try { refreshRanking(); } catch (err) { /* sem abas ainda */ }
}

function onEdit(e) {
  const sh = e.range.getSheet();
  const name = sh.getName();
  const r0 = e.range.getRow(), c0 = e.range.getColumn();
  const nr = e.range.getNumRows(), c1 = c0 + e.range.getNumColumns() - 1;
  const first = Math.max(r0, name === SH.res ? RANK_START : 2);
  const n = r0 + nr - first;
  if (n <= 0) return;
  const hits = col => c0 <= col && c1 >= col;

  const kind = Object.keys(RANKED).find(k => RANKED[k].sheet === name);
  if (kind) {
    const cfg = RANKED[kind];
    if (hits(cfg.titleCol)) {
      const titles = sh.getRange(first, cfg.titleCol, n, 1).getValues();
      sh.getRange(first, cfg.keyCol, n, 1).setValues(titles.map(t => [keyOf(kind, t[0])]));
    }
    if (c0 === 1) {
      syncDone(kind, sh.getRange(first, 1, n, 1).getValues(), sh.getRange(first, cfg.keyCol, n, 1).getValues());
    }
    refreshRanking();
  } else if (name === SH.res) {
    let changed = false;
    if (hits(RES.musDone)) {
      syncDone('mus', sh.getRange(first, RES.musDone, n, 1).getValues(), sh.getRange(first, RES.musKey, n, 1).getValues());
      changed = true;
    }
    if (hits(RES.santoDone)) {
      syncDone('santo', sh.getRange(first, RES.santoDone, n, 1).getValues(), sh.getRange(first, RES.santoKey, n, 1).getValues());
      changed = true;
    }
    if (changed) refreshRanking();
  } else if (name === SH.post && c0 === 1) {
    refreshRanking();
  }
}

// Marca/desmarca todas as linhas do mesmo item (repetidas).
function syncDone(kind, vals, keys) {
  const map = {};
  keys.forEach((k, i) => { if (k[0]) map[k[0]] = vals[i][0] === true; });
  if (!Object.keys(map).length) return;
  const cfg = RANKED[kind];
  const sh = ss().getSheetByName(cfg.sheet);
  const last = sh.getLastRow();
  if (last < 2) return;
  const a = sh.getRange(2, 1, last - 1, 1).getValues();
  const k = sh.getRange(2, cfg.keyCol, last - 1, 1).getValues();
  let changed = false;
  k.forEach((row, i) => {
    if (row[0] in map && a[i][0] !== map[row[0]]) { a[i][0] = map[row[0]]; changed = true; }
  });
  if (changed) sh.getRange(2, 1, last - 1, 1).setValues(a);
}

// ---------- rankings (sem repetição) ----------

// Agrupa as linhas de uma aba pela chave e grava a coluna "Pedidos".
function collect(kind) {
  const cfg = RANKED[kind];
  const sh = ss().getSheetByName(cfg.sheet);
  if (!sh) return { total: 0, list: [] };
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, cfg.keyCol).getValues() : [];
  const groups = {};
  const order = [];
  const keyFor = r => r[cfg.keyCol - 1] || keyOf(kind, r[cfg.titleCol - 1]);
  data.forEach(r => {
    const key = keyFor(r);
    if (!key) return;
    let g = groups[key];
    if (!g) { g = groups[key] = { key, name: String(r[cfg.titleCol - 1]), subs: {}, n: 0, done: false }; order.push(g); }
    g.n++;
    if (r[0] === true) g.done = true;
    const s = cfg.subCol ? String(r[cfg.subCol - 1]).trim() : '';
    if (s) g.subs[s] = (g.subs[s] || 0) + 1;
  });
  if (data.length) {
    sh.getRange(2, cfg.countCol, data.length, 1).setValues(data.map(r => {
      const g = groups[keyFor(r)];
      return [g ? g.n : ''];
    }));
  }
  const list = order.sort((a, b) => b.n - a.n || a.name.localeCompare(b.name, 'pt'));
  return { total: data.filter(r => String(r[cfg.titleCol - 1]).trim()).length, list };
}

function refreshRanking() {
  const res = ss().getSheetByName(SH.res);
  if (!res || !ss().getSheetByName(SH.mus)) return;
  const mus = collect('mus');
  const santo = collect('santo');
  writeCounters(res, mus, santo);

  const maxR = res.getMaxRows();
  const area = res.getRange(RANK_START, 2, maxR - RANK_START + 1, RES.cols - 1);
  area.breakApart().clearContent().clearDataValidations().setBackground(C.cream)
    .setBorder(false, false, false, false, false, false).setFontWeight('normal').setFontStyle('normal')
    .setFontColor(C.ink).setFontSize(10).setFontFamily(FONT).setHorizontalAlignment('left');

  const nm = mus.list.length, ns = santo.list.length;
  const need = Math.max(nm, ns, 1);
  if (RANK_START + need - 1 > maxR) res.insertRowsAfter(maxR, RANK_START + need - 1 - maxR);
  const empty = (col, msg) => res.getRange(RANK_START, col).setValue(msg)
    .setFontColor(C.muted).setFontStyle('italic');

  if (nm) {
    const top = g => Object.keys(g.subs).sort((a, b) => g.subs[b] - g.subs[a])[0] || '';
    const block = res.getRange(RANK_START, 2, nm, 6);
    block.setValues(mus.list.map(g => [g.n, g.name, '', top(g), '', false]));
    res.getRange(RANK_START, RES.musKey, nm, 1).setValues(mus.list.map(g => [g.key]));
    res.getRange(RANK_START, 3, nm, 2).mergeAcross();
    res.getRange(RANK_START, 5, nm, 2).mergeAcross();
    const cb = res.getRange(RANK_START, RES.musDone, nm, 1);
    cb.setDataValidation(checkboxRule()).setValues(mus.list.map(g => [g.done]));
    styleRank(res, block, nm, 2, C.gold);
    cb.setHorizontalAlignment('center');
    res.getRange(RANK_START, 5, nm, 1).setFontColor(C.muted);
  } else {
    empty(3, 'Nenhuma música ainda — compartilhe o link do site 🙂');
  }

  if (ns) {
    const block = res.getRange(RANK_START, 9, ns, 4);
    block.setValues(santo.list.map(g => [g.n, g.name, '', false]));
    res.getRange(RANK_START, RES.santoKey, ns, 1).setValues(santo.list.map(g => [g.key]));
    res.getRange(RANK_START, 10, ns, 2).mergeAcross();
    const cb = res.getRange(RANK_START, RES.santoDone, ns, 1);
    cb.setDataValidation(checkboxRule()).setValues(santo.list.map(g => [g.done]));
    styleRank(res, block, ns, 9, C.wine);
    cb.setHorizontalAlignment('center');
  } else {
    empty(10, 'Nenhum santo ainda 🙂');
  }
  res.setRowHeights(RANK_START, need, 30);
}

function styleRank(res, block, n, countCol, color) {
  block.setBackground(C.white).setVerticalAlignment('middle')
    .setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
  res.getRange(RANK_START, countCol, n, 1).setNumberFormat('0"×"').setFontWeight('bold')
    .setFontColor(color).setHorizontalAlignment('center').setFontSize(12);
}

function writeCounters(res, mus, santo) {
  const post = ss().getSheetByName(SH.post);
  const pl = post.getLastRow();
  const pv = pl > 1 ? post.getRange(2, 1, pl - 1, 3).getValues().filter(r => String(r[2]).trim()) : [];
  const mUni = mus.list.length, mDone = mus.list.filter(g => g.done).length;
  const sUni = santo.list.length, sDone = santo.list.filter(g => g.done).length;
  const pTot = pv.length, pDone = pv.filter(r => r[0] === true).length;
  res.getRange('B6:L6').setValues([[mus.total, mUni, mDone, mUni - mDone,
    santo.total, sUni, sDone, sUni - sDone, pTot, pDone, pTot - pDone]]);
  res.getRange('B7').setRichTextValue(bar(mDone, mUni, C.gold, 38));
  res.getRange('F7').setRichTextValue(bar(sDone, sUni, C.wine, 38));
  res.getRange('J7').setRichTextValue(bar(pDone, pTot, C.rose, 28));
  const pct = (a, b) => (b ? Math.round(100 * a / b) : 0) + '%';
  res.getRange('B8').setValue(`${pct(mDone, mUni)} das músicas já viraram edit`);
  res.getRange('F8').setValue(`${pct(sDone, sUni)} dos santos já foram postados`);
  res.getRange('J8').setValue(`${pct(pDone, pTot)} das ideias já viraram post`);
}

function bar(done, total, color, W) {
  const k = total ? Math.round(W * done / total) : 0;
  const s = '━'.repeat(W);
  const st = c => SpreadsheetApp.newTextStyle().setForegroundColor(c).setFontSize(12).build();
  const b = SpreadsheetApp.newRichTextValue().setText(s).setTextStyle(0, W, st(C.line));
  if (k) b.setTextStyle(0, k, st(color));
  return b.build();
}

// ---------- montagem das abas ----------

function setup() {
  const s = ss();
  const mus = s.getSheetByName(SH.mus) || buildMusicas(s);
  const santo = s.getSheetByName(SH.santo) || buildSantos(s);
  const post = s.getSheetByName(SH.post) || buildPosts(s);
  let res = s.getSheetByName(SH.res);
  // O Resumo só tem dados calculados: se for da versão sem santos, é refeito.
  if (res && !/SANTOS/.test(res.getRange('F4').getDisplayValue())) {
    s.setActiveSheet(mus);
    s.deleteSheet(res);
    res = null;
  }
  res = res || buildResumo(s);
  migrateOld(s, mus, post);
  styleProgress(res);
  [res, mus, santo, post].forEach((sh, i) => { s.setActiveSheet(sh); s.moveActiveSheet(i + 1); });
  s.setActiveSheet(res);
  refreshRanking();
  return 'ok';
}

function styleHeader(sh, n, widths) {
  sh.getRange(1, 1, 1, n).setBackground(C.ink).setFontColor(C.gold).setFontWeight('bold')
    .setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle');
  sh.setRowHeight(1, 38);
  sh.setFrozenRows(1);
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const extra = sh.getMaxColumns() - n;
  if (extra > 0) sh.deleteColumns(n + 1, extra);
}

function styleBody(sh, ncols, wrapCols, centerCols, doneRuleCols) {
  const rows = sh.getMaxRows();
  sh.getRange(2, 1, rows - 1, ncols).setFontFamily(FONT).setFontSize(10).setFontColor(C.ink)
    .setVerticalAlignment('middle');
  wrapCols.forEach(c => sh.getRange(2, c, rows - 1, 1).setWrap(true));
  centerCols.forEach(c => sh.getRange(2, c, rows - 1, 1).setHorizontalAlignment('center'));
  const band = sh.getRange(1, 1, rows, ncols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor(C.ink).setFirstRowColor(C.white).setSecondRowColor(C.cream);
  sh.getRange(1, 1, rows, ncols).createFilter();
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2')
    .setBackground(C.doneBg).setFontColor(C.muted).setStrikethrough(true)
    .setRanges([sh.getRange(2, 1, rows - 1, doneRuleCols)]).build();
}

function dupRule(sh, titleCol, countCol, color, bg) {
  const rows = sh.getMaxRows();
  const L = String.fromCharCode(64 + countCol);
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${L}2>1`)
    .setBackground(bg).setFontColor(color).setBold(true)
    .setRanges([sh.getRange(2, titleCol, rows - 1, 1), sh.getRange(2, countCol, rows - 1, 1)]).build();
}

function buildMusicas(s) {
  const sh = s.insertSheet(SH.mus);
  sh.getRange(1, 1, 1, MUS_HEAD.length).setValues([MUS_HEAD]);
  styleHeader(sh, MUS_HEAD.length, [56, 118, 220, 170, 78, 200, 170, 280, 150, 140, 80]);
  sh.hideColumns(RANKED.mus.keyCol);
  const done = styleBody(sh, 10, [3, 7, 8], [1, 2, 5], 10);
  sh.getRange(2, 5, sh.getMaxRows() - 1, 1).setNumberFormat('0"×"');
  sh.setConditionalFormatRules([done, dupRule(sh, 3, 5, C.rose, C.roseBg)]);
  sh.getRange('A1').setNote('Marque quando já tiver feito o edit. As repetidas são marcadas juntas.');
  sh.getRange('E1').setNote('Quantas vezes essa música foi sugerida. Em rosa = repetida.');
  sh.setTabColor(C.gold);
  return sh;
}

function buildSantos(s) {
  const sh = s.insertSheet(SH.santo);
  sh.getRange(1, 1, 1, SANTO_HEAD.length).setValues([SANTO_HEAD]);
  styleHeader(sh, SANTO_HEAD.length, [56, 118, 240, 78, 380, 150, 140, 80]);
  sh.hideColumns(RANKED.santo.keyCol);
  const done = styleBody(sh, 7, [3, 5], [1, 2, 4], 7);
  sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setNumberFormat('0"×"');
  sh.setConditionalFormatRules([done, dupRule(sh, 3, 4, C.wine, C.wineBg)]);
  sh.getRange('A1').setNote('Marque quando o santo já tiver saído no Santo da Semana. As repetidas são marcadas juntas.');
  sh.getRange('D1').setNote('Quantas vezes esse santo foi sugerido. Em vinho = repetido.');
  sh.setTabColor(C.wine);
  return sh;
}

function buildPosts(s) {
  const sh = s.insertSheet(SH.post);
  sh.getRange(1, 1, 1, POST_HEAD.length).setValues([POST_HEAD]);
  styleHeader(sh, POST_HEAD.length, [56, 118, 520, 160, 150]);
  const done = styleBody(sh, 5, [3], [1, 2], 5);
  sh.setConditionalFormatRules([done]);
  sh.getRange('A1').setNote('Marque quando já tiver feito o post.');
  sh.setTabColor(C.rose);
  return sh;
}

function styleProgress(sh) {
  ['B7:E7', 'F7:I7', 'J7:L7', 'B8:E8', 'F8:I8', 'J8:L8']
    .forEach(a => sh.getRange(a).breakApart().merge().clearContent());
  sh.getRange('B7:L7').setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange('B8:L8').setFontColor(C.muted).setFontSize(9).setFontFamily(FONT);
  sh.setRowHeight(7, 22);
}

function buildResumo(s) {
  const sh = s.insertSheet(SH.res, 0);
  sh.setHiddenGridlines(true);
  const extra = sh.getMaxColumns() - RES.cols;
  if (extra > 0) sh.deleteColumns(RES.cols + 1, extra);
  const rows = sh.getMaxRows();
  sh.getRange(1, 1, rows, RES.cols).setBackground(C.cream).setFontFamily(FONT).setFontColor(C.ink)
    .setVerticalAlignment('middle');
  sh.setColumnWidth(1, 28);
  for (let c = 2; c <= 12; c++) sh.setColumnWidth(c, 118);
  sh.setColumnWidth(13, 28);
  sh.hideColumns(RES.musKey, 2);
  sh.setTabColor(C.ink);

  sh.setRowHeight(1, 16);
  sh.getRange('B2').setValue('valei-me').setFontFamily('Playfair Display').setFontSize(30)
    .setFontWeight('bold');
  sh.setRowHeight(2, 54);
  sh.getRange('B3').setValue('Sugestões que chegam pelo site  ·  valeime.netlify.app')
    .setFontColor(C.muted).setFontSize(10);
  sh.setRowHeight(3, 24);
  sh.setRowHeight(4, 34);

  [['B4:E4', '🎵  MÚSICAS', C.gold], ['F4:I4', '✦  SANTOS', C.wine], ['J4:L4', '✝  POSTS', C.rose]]
    .forEach(([a, t, c]) => sh.getRange(a).merge().setValue(t).setFontColor(c).setFontWeight('bold').setFontSize(11));

  const cards = [
    ['B', 'Sugestões', C.ink], ['C', 'Músicas diferentes', C.gold], ['D', 'Já feitas', C.gold], ['E', 'Faltam', C.ink],
    ['F', 'Sugestões', C.ink], ['G', 'Santos diferentes', C.wine], ['H', 'Já feitos', C.wine], ['I', 'Faltam', C.ink],
    ['J', 'Ideias', C.rose], ['K', 'Já feitas', C.rose], ['L', 'Faltam', C.ink],
  ];
  cards.forEach(([col, label, color]) => {
    sh.getRange(`${col}5`).setValue(label.toUpperCase()).setFontSize(8).setFontColor(C.muted)
      .setFontWeight('bold').setBackground(C.white).setHorizontalAlignment('center')
      .setVerticalAlignment('bottom');
    sh.getRange(`${col}6`).setValue(0).setFontSize(28).setFontFamily('Playfair Display')
      .setFontWeight('bold').setFontColor(color).setBackground(C.white).setHorizontalAlignment('center');
    sh.getRange(`${col}5:${col}6`).setBorder(true, true, true, true, null, null, C.line,
      SpreadsheetApp.BorderStyle.SOLID);
  });
  ['F4:F6', 'J4:J6'].forEach(a => sh.getRange(a)
    .setBorder(null, true, null, null, null, null, C.ink, SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
  sh.setRowHeight(5, 26);
  sh.setRowHeight(6, 58);

  styleProgress(sh);
  sh.setRowHeight(8, 22);

  sh.getRange('B9').setValue('🏆  Músicas mais pedidas (sem repetidas)').setFontFamily('Playfair Display')
    .setFontSize(14).setFontWeight('bold');
  sh.getRange('I9').setValue('✦  Santos mais pedidos').setFontFamily('Playfair Display')
    .setFontSize(14).setFontWeight('bold');
  sh.setRowHeight(9, 46);

  const headM = sh.getRange('B10:G10');
  headM.setValues([['Pedidos', 'Música', '', 'Artista', '', 'Feito']]);
  sh.getRange('C10:D10').merge();
  sh.getRange('E10:F10').merge();
  const headS = sh.getRange('I10:L10');
  headS.setValues([['Pedidos', 'Santo', '', 'Feito']]);
  sh.getRange('J10:K10').merge();
  [headM, headS].forEach(h => h.setBackground(C.ink).setFontColor(C.gold).setFontWeight('bold').setFontSize(10));
  ['B10', 'G10', 'I10', 'L10'].forEach(a => sh.getRange(a).setHorizontalAlignment('center'));
  sh.getRange('G10').setNote('Marque aqui ou na aba Músicas — as repetidas são marcadas juntas.');
  sh.getRange('L10').setNote('Marque aqui ou na aba Santos — os repetidos são marcados juntos.');
  sh.setRowHeight(10, 32);
  sh.setFrozenRows(0);

  const n = rows - RANK_START + 1;
  const rule = (col, range) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${col}${RANK_START}`)
    .setFontColor(C.muted).setStrikethrough(true).setBackground(C.doneBg)
    .setRanges([range]).build();
  sh.setConditionalFormatRules([
    rule('G', sh.getRange(RANK_START, 2, n, 6)),
    rule('L', sh.getRange(RANK_START, 9, n, 4)),
  ]);
  return sh;
}

// Move linhas da aba antiga (Página1) para as novas abas e apaga a antiga.
function migrateOld(s, mus, post) {
  s.getSheets().forEach(sh => {
    if (Object.values(SH).includes(sh.getName())) return;
    const vals = sh.getDataRange().getValues();
    if (!vals.length || vals[0][0] !== 'Data' || vals[0][1] !== 'Tipo') return;
    vals.slice(1).forEach(r => {
      const [date, tipo, musica, artista, link, trecho, ideia, nome, insta] = r;
      if (tipo === 'Teste' || /pode apagar/i.test(musica)) return;
      if (tipo === 'Ideia de post') {
        const row = nextRow(post);
        post.getRange(row, 1, 1, 5).setValues([[false, date, ideia, nome, insta]]);
        post.getRange(row, 1).setDataValidation(checkboxRule());
      } else if (musica) {
        const row = nextRow(mus);
        mus.getRange(row, 1, 1, 11).setValues([[false, date, musica, artista, '', link, trecho, ideia, nome, insta, norm(musica)]]);
        mus.getRange(row, 1).setDataValidation(checkboxRule());
      }
    });
    s.deleteSheet(sh);
  });
}

function limparTestes() {
  [[SH.mus, 9], [SH.santo, 6], [SH.post, 4]].forEach(([name, col]) => {
    const sh = ss().getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const v = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
    for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === 'teste-copilot') sh.deleteRow(i + 2);
  });
  refreshRanking();
  return 'ok';
}
