/* Сверка актов по аренде. Всё считается в браузере: файлы никуда не отправляются.
 * Читает акт сверки 1С (xlsx / xls / csv / pdf с текстовым слоем), сопоставляет операции сторон,
 * разбирает оплаты по месяцам (нет оплаты, недоплата, переплата, сдвиг), учитывает обеспечительный
 * платёж и его доплату при индексации и выгружает разбор в Excel.
 * Логика повторяет навык Claude «akt-sverki» (scripts/reconcile.py). */
(() => {
  'use strict';
  const LIB = {
    xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    exceljs: 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js',
    pdf: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
    pdfWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  };
  const loaded = {};
  const loadScript = src => loaded[src] || (loaded[src] = new Promise((ok, fail) => {
    const s = document.createElement('script');
    s.src = src; s.onload = ok;
    s.onerror = () => { delete loaded[src]; s.remove(); fail(new Error('Не удалось загрузить библиотеку для чтения файлов. Проверьте интернет.')); };
    document.head.appendChild(s);
  }));

  const EPS = 0.005, ACC = 'acc', PAY = 'pay';
  const r2 = x => Math.round(x * 100) / 100;
  const fmt = x => (Math.abs(x) < EPS ? 0 : x).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const MONTHS = 'янв фев мар апр май июн июл авг сен окт ноя дек'.split(' ');
  const mlabel = k => `${MONTHS[k % 12]} ${Math.floor(k / 12)}`; // k = год * 12 + месяц (0..11)
  const esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

  const ROLES = {
    tenant: {
      cp: 'арендодатель', gen: 'арендодателя', dat: 'арендодателю', debt: 'Наш долг арендодателю',
      over: 'Мы заплатили больше, чем начислено. Проверить ставку и дату индексации по договору. Если начисление верное — договориться о зачёте в счёт будущих платежей или о возврате.',
      under: 'Мы заплатили меньше, чем начислено. Проверить, верна ли ставка у арендодателя; если да — доплатить.',
      nopay: 'Нет нашей оплаты за месяц. Найти платёж в выписке (мог уйти с другим назначением) или оплатить.',
    },
    landlord: {
      cp: 'арендатор', gen: 'арендатора', dat: 'арендатору', debt: 'Долг арендатора нам',
      over: 'Арендатор заплатил больше, чем мы начислили. Проверить, не забыли ли мы начислить (индексация, коммуналка). Если нет — зачесть в счёт будущих платежей или вернуть по заявлению арендатора.',
      under: 'Арендатор заплатил меньше, чем начислено. Проверить начисление; если оно верное — направить арендатору напоминание о доплате.',
      nopay: 'Нет оплаты за месяц. Найти платёж в выписке (мог прийти с другим назначением) или направить арендатору напоминание / претензию.',
    },
  };

  // ---------- даты и числа ----------
  function mkDate(y, m, d) {
    const t = new Date(Date.UTC(y, m - 1, d));
    if (t.getUTCMonth() !== m - 1 || t.getUTCDate() !== d) return null;
    return { y, m, d, n: Math.round(t.getTime() / 864e5), k: y * 12 + m - 1 };
  }
  function parseDate(v) {
    if (v instanceof Date && !isNaN(v)) {
      const t = new Date(v.getTime() + 12 * 36e5); // Excel-даты иногда смещены на секунды/часы
      return mkDate(t.getFullYear(), t.getMonth() + 1, t.getDate());
    }
    if (v == null) return null;
    const m = String(v).match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
    if (!m) return null;
    let y = +m[3]; if (y < 100) y += 2000;
    return mkDate(y, +m[2], +m[1]);
  }
  const dstr = t => `${String(t.d).padStart(2, '0')}.${String(t.m).padStart(2, '0')}.${t.y}`;
  const xlDate = t => new Date(Date.UTC(t.y, t.m - 1, t.d));
  function parseNum(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    let s = String(v).trim().replace(/[\s  ]/g, '').replace(',', '.');
    const neg = /^\(.*\)$/.test(s); s = s.replace(/[()]/g, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return 0;
    return neg ? -parseFloat(s) : parseFloat(s);
  }
  function text(v) {
    if (v == null) return '';
    if (v instanceof Date) { const d = parseDate(v); return d ? dstr(d) : ''; }
    return String(v).trim();
  }

  // ---------- чтение файлов в таблицу ----------
  function parseCsv(str) {
    const first = str.split(/\r?\n/, 1)[0];
    const delim = [';', '\t', ','].sort((a, b) => first.split(b).length - first.split(a).length)[0];
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < str.length; i++) {
      const ch = str[i];
      if (q) { if (ch === '"') { if (str[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; continue; }
      if (ch === '"') q = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n') { row.push(cell.replace(/\r$/, '')); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }
  async function readGrid(file) {
    const ext = ((file.name.match(/\.([^.]+)$/) || [])[1] || '').toLowerCase();
    const buf = await file.arrayBuffer();
    if (ext === 'pdf') return pdfGrid(buf);
    if (ext === 'csv' || ext === 'txt') {
      let s;
      try { s = new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch (e) { s = new TextDecoder('windows-1251').decode(buf); }
      return parseCsv(s.replace(/^﻿/, ''));
    }
    if (['xlsx', 'xlsm', 'xls', 'ods'].includes(ext)) {
      await loadScript(LIB.xlsx);
      const wb = window.XLSX.read(buf, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    }
    throw new Error(`Файл «${file.name}»: нужен Excel (.xlsx, .xls), CSV или PDF из 1С.`);
  }

  // PDF из 1С: текст лежит кусками с координатами. Собираем строки по высоте, а колонки — по
  // заголовку таблицы («Дата», «Документ», «Дебет», «Кредит»): суммы относим к ближайшей колонке
  // Дебет/Кредит, остальной текст — к ближайшей текстовой колонке.
  async function pdfGrid(buf) {
    await loadScript(LIB.pdf);
    const lib = window.pdfjsLib;
    lib.GlobalWorkerOptions.workerSrc = LIB.pdfWorker;
    const doc = await lib.getDocument({ data: new Uint8Array(buf) }).promise;
    const lines = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const tc = await (await doc.getPage(p)).getTextContent();
      const its = tc.items.filter(i => i.str && i.str.trim())
        .map(i => ({ s: i.str.trim(), x: i.transform[4], y: i.transform[5], w: i.width }))
        .sort((a, b) => b.y - a.y || a.x - b.x);
      let cur = null;
      for (const it of its) {
        if (!cur || Math.abs(cur.y - it.y) > 2.5) { cur = { y: it.y, items: [] }; lines.push(cur); }
        cur.items.push(it);
      }
    }
    if (!lines.length) throw new Error('В PDF нет текста — похоже, это скан. Сохраните акт из 1С в Excel или отправьте скан Claude.');
    const hi = lines.findIndex(ln => ln.items.some(t => /^дебет$/i.test(t.s)) && ln.items.some(t => /^кредит$/i.test(t.s)));
    if (hi < 0) throw new Error('В PDF не найдена таблица акта с колонками «Дебет» и «Кредит».');
    const anchors = lines[hi].items.map(t => ({ s: t.s, c: t.x + t.w / 2, num: /^(дебет|кредит)$/i.test(t.s) })).sort((a, b) => a.c - b.c);
    const isNum = s => /^-?\d[\d\s ]*,\d{2}$/.test(s);
    return lines.map((ln, i) => {
      if (i === hi) return anchors.map(a => a.s);
      const row = Array(anchors.length).fill(null);
      for (const t of ln.items) {
        const c = t.x + t.w / 2, num = isNum(t.s);
        let best = -1, bd = Infinity;
        anchors.forEach((a, k) => { if (a.num === num && Math.abs(a.c - c) < bd) { bd = Math.abs(a.c - c); best = k; } });
        if (best >= 0) row[best] = row[best] ? row[best] + ' ' + t.s : t.s;
      }
      return row;
    });
  }

  // ---------- разбор акта ----------
  function findBlocks(grid) {
    for (let hr = 0; hr < Math.min(grid.length, 80); hr++) {
      const cells = (grid[hr] || []).map(c => text(c).toLowerCase());
      const debits = cells.flatMap((c, i) => c.startsWith('дебет') || c === 'дт' ? [i] : []);
      const credits = cells.flatMap((c, i) => c.startsWith('кредит') || c === 'кт' ? [i] : []);
      if (!debits.length || !credits.length) continue;
      const blocks = []; let prevEnd = -1;
      for (const d of debits) {
        const k = credits.find(c => c > d); if (k == null) continue;
        let date = null, doc = null;
        for (const r of [hr - 1, hr]) {
          if (r < 0) continue;
          for (let i = prevEnd + 1; i < d; i++) {
            const t = text((grid[r] || [])[i]).toLowerCase();
            if (t.includes('дата') && date == null) date = i;
            if (doc == null && /документ|содержан|операц|наименован/.test(t)) doc = i;
          }
        }
        blocks.push({ date: date ?? prevEnd + 1, doc, debit: d, credit: k, start: prevEnd + 1 });
        prevEnd = k;
      }
      if (blocks.length) return { hr, blocks };
    }
    throw new Error('Не нашёл в файле таблицу акта: нет заголовка с колонками «Дебет» и «Кредит».');
  }
  function blockLabel(grid, hr, b) {
    for (let r = Math.max(0, hr - 4); r < hr; r++) {
      for (let i = b.start; i <= b.credit; i++) {
        const t = text((grid[r] || [])[i]);
        if (/^по данным/i.test(t)) return t.replace(/,?\s*руб\.?\s*$/i, '').replace(/^по данным\s+/i, '');
      }
    }
    return '';
  }
  // accInCredit: у арендатора начисление в Кредите, у арендодателя — в Дебете
  function readBlock(grid, hr, b, side, accInCredit) {
    const cell = (row, i) => (i != null && row && i < row.length ? row[i] : null);
    const sign = accInCredit ? 1 : -1;
    const rows = []; let opening = 0, closing = null;
    for (let r = hr + 1; r < grid.length; r++) {
      const row = grid[r];
      const dv = parseNum(cell(row, b.debit)), kv = parseNum(cell(row, b.credit));
      const span = []; for (let i = b.start; i < b.debit; i++) { const t = text(cell(row, i)); if (t) span.push(t); }
      const lab = span.join(' ').toLowerCase();
      const bal = sign * (kv - dv);
      if (lab.includes('сальдо') && /начал/.test(lab)) { opening = bal; continue; }
      if (lab.includes('сальдо') && /конеч|на конец/.test(lab)) { closing = bal; break; }
      if (lab.includes('оборот')) continue;
      let dt = parseDate(cell(row, b.date));
      const docCols = b.doc == null ? Array.from({ length: Math.max(0, b.debit - b.date - 1) }, (_, j) => b.date + 1 + j) : [b.doc];
      const doc = docCols.map(i => text(cell(row, i))).filter(Boolean).join(' ');
      if (!dt) dt = parseDate(doc);
      if (!dt || (Math.abs(dv) < EPS && Math.abs(kv) < EPS)) continue;
      const [acc, pay] = accInCredit ? [kv, dv] : [dv, kv];
      if (Math.abs(acc) >= EPS) rows.push(mkRow(side, r + 1, dt, doc, ACC, r2(acc)));
      if (Math.abs(pay) >= EPS) rows.push(mkRow(side, r + 1, dt, doc, PAY, r2(pay)));
    }
    return { rows, opening: r2(opening), closingStated: closing == null ? null : r2(closing) };
  }
  function mkRow(side, line, dt, doc, kind, amt) {
    const m = doc.match(/№\s*([^\s(),]+)/) || doc.match(/\(([^\s()]+)\s+от\s+\d/);
    return { side, line, dt, doc, kind, amt, docNo: m ? m[1].replace(/^0+/, '') : '', group: 0, status: '', note: '' };
  }
  const effect = r => r.kind === ACC ? r.amt : -r.amt;
  const closingOf = s => r2(s.opening + s.rows.reduce((a, r) => a + effect(r), 0));

  // ---------- сопоставление сторон ----------
  function reasons(R) {
    return {
      ok: ['Совпало', ''], shift: ['Совпало, даты разные', 'Операция учтена разными датами внутри периода. На сальдо не влияет.'],
      period: ['Разные месяцы', 'Стороны отнесли операцию к разным месяцам. Проверить дату в первичном документе.'],
      split: ['Одна = несколько', 'Одна операция у одной стороны разбита на несколько у другой. Расхождения нет.'],
      round: ['Копейки', 'Разница в копейках/округлении. Согласовать, по какой сумме закрыть.'],
      vat: ['Разница на НДС', 'Одна сторона учла сумму с НДС, другая без. Сверить с УПД / счётом-фактурой.'],
      amount: ['Разные суммы', 'Один документ, разные суммы. Сверить с первичным документом и договором.'],
      their_acc: [`Начисление только у ${R.gen}`, `Найти первичный документ (акт/УПД) и проверить по договору. Если начисление верное — провести у нас; если нет — попросить ${R.gen} исправить.`],
      their_pay: [`Оплата только у ${R.gen}`, 'Проверить банковскую выписку: платёж мог пройти с другим назначением или на другое юрлицо.'],
      our_acc: ['Начисление только у нас', `Уточнить у ${R.gen}, получил ли он документ; возможно, он учёл его в другом периоде или по другому договору.`],
      our_pay: ['Оплата только у нас', `Отправить ${R.dat} подтверждение платежа (платёжное поручение: дата, номер, сумма, назначение).`],
    };
  }
  const KEYWORDS = [
    [/пени|неустой|штраф/i, 'похоже на пени/неустойку — проверить основание по договору'],
    [/обеспечит|депозит|гарантийн|залог/i, 'похоже на обеспечительный платёж — проверить, где его учли стороны'],
    [/электр|коммун|вод[аы]|тепл|отоплен|эксплуат|возмещ|вывоз|мусор/i, 'похоже на коммунальные/возмещаемые — нужны показания и расчёт'],
    [/возврат/i, 'возврат денег — проверить выписку'],
    [/корректир|исправлен|сторно/i, 'корректировочный документ — проверить, проведён ли он у обеих сторон'],
    [/индекс/i, 'индексация ставки — проверить дату и процент по договору'],
  ];
  const hint = doc => (KEYWORDS.find(([re]) => re.test(doc)) || [, ''])[1];

  function match(our, their, R, maxDays = 45, rounding = 1) {
    let g = 1;
    const free = rows => rows.filter(r => !r.group);
    const link = (a, b, status, note = '') => { const id = g++; [...a, ...b].forEach(r => { r.group = id; r.status = status; r.note = note; }); };
    const days = (o, t) => Math.abs(o.dt.n - t.dt.n);
    const same = (o, t) => Math.abs(o.amt - t.amt) < EPS;
    const nums = (o, t) => o.docNo && o.docNo === t.docNo;
    const amounts = (o, t) => `у нас ${fmt(o.amt)}, у ${R.gen} ${fmt(t.amt)}`;
    const pairwise = (cond, statusFn) => {
      for (const o of free(our.rows)) {
        const c = free(their.rows).filter(t => t.kind === o.kind && cond(o, t));
        if (!c.length) continue;
        const t = c.reduce((a, b) => days(o, b) < days(o, a) ? b : a);
        const [st, note] = statusFn(o, t); link([o], [t], st, note);
      }
    };
    const dateStatus = (o, t) => {
      if (o.dt.n === t.dt.n) return ['ok', ''];
      const note = `у нас ${dstr(o.dt)}, у ${R.gen} ${dstr(t.dt)}`;
      return [o.dt.k !== t.dt.k ? 'period' : 'shift', note];
    };
    pairwise((o, t) => same(o, t) && o.dt.n === t.dt.n, dateStatus);
    pairwise((o, t) => same(o, t) && nums(o, t) && days(o, t) <= 120, dateStatus);
    pairwise((o, t) => same(o, t) && days(o, t) <= maxDays, dateStatus);
    const split = (bigRows, smallRows) => {
      for (const x of free(bigRows)) {
        const c = free(smallRows).filter(r => r.kind === x.kind && r.amt > 0 && r.amt < x.amt && Math.abs(r.dt.n - x.dt.n) <= maxDays)
          .sort((a, b) => Math.abs(a.dt.n - x.dt.n) - Math.abs(b.dt.n - x.dt.n)).slice(0, 12);
        let found = null;
        const rec = (start, n, acc, sum) => {
          if (found) return;
          if (acc.length === n) { if (Math.abs(sum - x.amt) < EPS) found = [...acc]; return; }
          for (let i = start; i < c.length && !found; i++) { acc.push(c[i]); rec(i + 1, n, acc, sum + c[i].amt); acc.pop(); }
        };
        for (const n of [2, 3, 4]) { rec(0, n, [], 0); if (found) break; }
        if (found) {
          const note = `${fmt(x.amt)} = ${found.map(r => fmt(r.amt)).join(' + ')}`;
          x.side === 'our' ? link([x], found, 'split', note) : link(found, [x], 'split', note);
        }
      }
    };
    split(our.rows, their.rows); split(their.rows, our.rows);
    pairwise((o, t) => Math.abs(o.amt - t.amt) > 0 && Math.abs(o.amt - t.amt) <= rounding && days(o, t) <= 10, (o, t) => ['round', amounts(o, t)]);
    pairwise((o, t) => {
      if (days(o, t) > 10 || Math.min(o.amt, t.amt) <= 0) return false;
      const q = Math.max(o.amt, t.amt) / Math.min(o.amt, t.amt);
      return [1.2, 1.05, 1.07, 1.22].some(k => Math.abs(q - k) < 0.0005);
    }, (o, t) => ['vat', amounts(o, t)]);
    pairwise((o, t) => (nums(o, t) && days(o, t) <= 31) || o.dt.n === t.dt.n,
      (o, t) => ['amount', `${amounts(o, t)}, разница ${fmt(t.amt - o.amt)}`]);
    free(their.rows).forEach(r => link([], [r], 'their_' + r.kind, hint(r.doc)));
    free(our.rows).forEach(r => link([r], [], 'our_' + r.kind, hint(r.doc)));

    const groups = new Map();
    [...our.rows, ...their.rows].forEach(r => {
      if (!groups.has(r.group)) groups.set(r.group, { our: [], their: [], status: r.status, note: r.note });
      groups.get(r.group)[r.side].push(r);
    });
    return [...groups.values()].map(gr => ({
      ...gr, impact: r2(gr.their.reduce((a, r) => a + effect(r), 0) - gr.our.reduce((a, r) => a + effect(r), 0)),
      first: Math.min(...[...gr.our, ...gr.their].map(r => r.dt.n)),
    })).sort((a, b) => a.first - b.first);
  }

  // ---------- помесячный разбор и обеспечительный ----------
  const MONTH_STATUS = { ok: 'Оплачено точно', shift: 'Оплата/начисление в другом месяце', over: 'Переплата', under: 'Недоплата', nopay: 'Нет оплаты', empty: '—', deposit: 'Обеспечительный платёж' };
  function regularRate(mo) {
    if (!mo.accs.length) return null;
    const cnt = new Map(); mo.accs.forEach(a => cnt.set(a, (cnt.get(a) || 0) + 1));
    return [...cnt.entries()].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0][0];
  }
  function rateSegments(ms) {
    const segs = [];
    ms.forEach((mo, i) => {
      const r = regularRate(mo); if (r == null) return;
      if (segs.length && Math.abs(segs[segs.length - 1][0] - r) < EPS) segs[segs.length - 1][2] = i; else segs.push([r, i, i]);
    });
    const count = s => ms.slice(s[1], s[2] + 1).filter(m => m.accs.length).length;
    const merged = [];
    segs.forEach((s, k) => {
      if (count(s) < 2 && k !== segs.length - 1) return; // разовое отклонение — не индексация
      const last = merged[merged.length - 1];
      if (last && Math.abs(last[0] - s[0]) < EPS) last[2] = s[2]; else merged.push([...s]);
    });
    return merged;
  }
  function analyzeMonths(side, window = 2, depMult = 1) {
    const ks = side.rows.map(r => r.dt.k);
    const k0 = Math.min(...ks), k1 = Math.max(...ks);
    const ms = [];
    for (let k = k0; k <= k1; k++) ms.push({ k, acc: 0, pay: 0, accs: [], pays: [], res: 0, dep: 0, notes: [], status: '', balance: 0 });
    side.rows.forEach(r => {
      const mo = ms[r.dt.k - k0];
      if (r.kind === ACC) { mo.acc = r2(mo.acc + r.amt); mo.accs.push(r.amt); } else { mo.pay = r2(mo.pay + r.amt); mo.pays.push(r.amt); }
    });
    ms.forEach(mo => { mo.diff = r2(mo.pay - mo.acc); mo.res = mo.diff; });
    // оплата «за два месяца разом» и пустой соседний месяц гасят друг друга
    ms.forEach((mo, i) => {
      if (mo.res >= -EPS) return;
      const c = [];
      for (let j = Math.max(0, i - window); j <= Math.min(ms.length - 1, i + window); j++) if (j !== i && ms[j].res > EPS) c.push(j);
      c.sort((a, b) => Math.abs(a - i) - Math.abs(b - i) || a - b);
      for (const j of c) {
        const t = r2(Math.min(ms[j].res, -mo.res)), o = ms[j];
        o.res = r2(o.res - t); mo.res = r2(mo.res + t);
        if (o.acc < EPS) { o.notes.push(`нет начисления — похоже, оно проведено в ${mlabel(mo.k)}`); mo.notes.push(`в т.ч. начисление за ${mlabel(o.k)} (${fmt(t)})`); }
        else if (j < i) { o.notes.push(`${fmt(t)} — оплата вперёд за ${mlabel(mo.k)}`); mo.notes.push(`${fmt(t)} оплачено заранее, в ${mlabel(o.k)}`); }
        else { o.notes.push(`${fmt(t)} — оплата за ${mlabel(mo.k)} с опозданием`); mo.notes.push(`${fmt(t)} оплачено позже, в ${mlabel(o.k)}`); }
        if (mo.res > -EPS) break;
      }
    });
    const dep = attributeDeposit(ms, side, depMult);
    let bal = side.opening;
    ms.forEach((mo, i) => {
      bal = r2(bal + mo.acc - mo.pay); mo.balance = bal;
      if (mo.acc < EPS && mo.pay < EPS) mo.status = 'empty';
      else if (Math.abs(mo.res) < EPS) mo.status = mo.dep > EPS ? 'deposit' : Math.abs(mo.diff) < EPS ? 'ok' : 'shift';
      else if (mo.res > 0) mo.status = 'over';
      else mo.status = mo.pay < EPS ? 'nopay' : 'under';
      if (['over', 'under', 'nopay'].includes(mo.status)) {
        if (mo.acc < EPS) mo.notes.unshift('оплата без начисления в этом месяце');
        else if (mo.pays.length === 1 && mo.accs.length === 1) mo.notes.unshift(`оплачено ${fmt(mo.pay)} при начислении ${fmt(mo.acc)}`);
        if (mo.status === 'nopay' && i === ms.length - 1) mo.notes.push('последний месяц периода — проверьте, наступил ли срок оплаты');
      }
    });
    return { ms, dep, periods: monthPeriods(ms) };
  }
  // переплату в начале договора относим к обеспечительному (= mult × первая ставка),
  // переплату рядом с индексацией (−3…+2 мес.) — к его доплате (= mult × рост ставки)
  function attributeDeposit(ms, side, mult) {
    const segs = rateSegments(ms);
    const info = { events: [], expected: 0, found: 0, mult, note: '', finalRate: null, hidden: false };
    if (!(mult > 0) || !segs.length) return info;
    info.finalRate = segs[segs.length - 1][0];
    info.expected = r2(info.finalRate * mult);
    const take = (lo, hi, center, need, label) => {
      let got = 0;
      const idx = [];
      for (let j = Math.max(0, lo); j <= Math.min(ms.length - 1, hi); j++) if (ms[j].res > EPS) idx.push(j);
      idx.sort((a, b) => Math.abs(a - center) - Math.abs(b - center) || a - b);
      for (const j of idx) {
        if (need - got < EPS) break;
        const t = r2(Math.min(ms[j].res, need - got));
        ms[j].res = r2(ms[j].res - t); ms[j].dep = r2(ms[j].dep + t);
        ms[j].notes.push(`${fmt(t)} — ${label}`); got = r2(got + t);
      }
      return got;
    };
    const first = segs[0];
    if (Math.abs(side.opening) < EPS) {
      const exp = r2(first[0] * mult);
      info.events.push({ what: `Первоначальный (ставка ${fmt(first[0])})`, when: mlabel(ms[first[1]].k), expected: exp, found: take(0, first[1] + 2, first[1], exp, 'обеспечительный платёж') });
    } else if (side.opening < 0) {
      info.note = `На начало периода аванс арендатора ${fmt(-side.opening)}. Обеспечительный, внесённый раньше, сидит в нём; при ставке ${fmt(first[0])} он должен быть ${fmt(first[0] * mult)}.`;
    } else {
      info.hidden = true;
      info.note = `На начало периода не аванс, а долг арендатора ${fmt(side.opening)} — обеспечительного в акте не видно. Скорее всего, он учтён по отдельному договору или счёту.`;
    }
    for (let s = 1; s < segs.length; s++) {
      const prev = segs[s - 1], cur = segs[s];
      const delta = r2((cur[0] - prev[0]) * mult);
      if (delta <= EPS) continue;
      info.events.push({ what: `Доплата при индексации ${fmt(prev[0])} → ${fmt(cur[0])}`, when: mlabel(ms[cur[1]].k), expected: delta,
        found: take(cur[1] - 3, cur[1] + 2, cur[1], delta, `доплата обеспечительного при индексации ${fmt(prev[0])} → ${fmt(cur[0])}`) });
    }
    info.found = r2(info.events.reduce((a, e) => a + e.found, 0));
    if (info.events.length && info.found < EPS) {
      info.hidden = true;
      info.note = 'В оплатах обеспечительный не виден — возможно, он учтён по отдельному договору или счёту.';
    }
    return info;
  }
  function monthPeriods(ms) {
    const out = []; let cur = null;
    ms.forEach(mo => {
      if (['over', 'under', 'nopay'].includes(mo.status)) {
        if (cur && cur.status === mo.status) cur.months.push(mo); else { cur = { status: mo.status, months: [mo] }; out.push(cur); }
      } else if (!['shift', 'deposit'].includes(mo.status)) cur = null;
    });
    const lastK = ms.length ? ms[ms.length - 1].k : null;
    out.forEach(p => {
      const m = p.months;
      p.total = r2(m.reduce((a, x) => a + x.res, 0));
      p.label = m.length === 1 ? mlabel(m[0].k) : `${mlabel(m[0].k)} – ${mlabel(m[m.length - 1].k)}`;
      const parts = [];
      if (m.length > 1) parts.push(new Set(m.map(x => x.res)).size === 1 ? `по ${fmt(Math.abs(m[0].res))} × ${m.length} мес.` : m.map(x => fmt(Math.abs(x.res))).join(' + '));
      const pairs = new Set(m.filter(x => x.pays.length === 1 && x.accs.length === 1).map(x => `${x.pay}|${x.acc}`));
      if (pairs.size === 1) { const [pay, acc] = [...pairs][0].split('|').map(Number); parts.push(`оплачено ${fmt(pay)} при начислении ${fmt(acc)}`); }
      else if (m.length === 1) {
        const x = m[0], moved = r2(x.diff - x.res - x.dep);
        let t = `оплачено ${fmt(x.pay)} при начислении ${fmt(x.acc)}`;
        if (Math.abs(moved) > EPS) t += `, из них ${fmt(Math.abs(moved))} — в счёт соседних месяцев`;
        if (x.dep > EPS) t += `, ${fmt(x.dep)} — обеспечительный`;
        parts.push(t);
      }
      if (m[m.length - 1].k === lastK && p.status === 'over') parts.push('последний месяц периода — возможно, это оплата вперёд за следующий месяц');
      if (m[m.length - 1].k === lastK && p.status === 'nopay') parts.push('последний месяц периода — проверьте, наступил ли срок оплаты');
      p.detail = parts.join('; ');
    });
    return out;
  }
  function depositLines(side, dep, periods) {
    if (!dep.finalRate) return [];
    const lines = [];
    const advance = r2(-closingOf(side));
    const rest = r2(periods.reduce((a, p) => a + p.total, 0));
    lines.push([`Обеспечительный должен быть ${fmt(dep.expected)} (${dep.mult} × текущая ставка ${fmt(dep.finalRate)}).`, true]);
    if (dep.events.length) {
      const short = r2(dep.events.reduce((a, e) => a + Math.max(0, e.expected - e.found), 0));
      lines.push([`Найдено в оплатах за период: ${fmt(dep.found)}${short > EPS ? `, не хватает ${fmt(short)}` : ''}.`, false]);
    }
    if (dep.note) lines.push([dep.note, false]);
    const parts = [`обеспечительный ${fmt(dep.found)}`, `переплаты/недоплаты по месяцам ${fmt(rest)}`];
    if (Math.abs(side.opening) > EPS) parts.unshift(`аванс на начало ${fmt(-side.opening)}`);
    lines.push([`Аванс арендатора на конец периода ${fmt(advance)} = ${parts.join(' + ')}.`, false]);
    const excess = r2(advance - dep.expected);
    if (dep.hidden) lines.push(['Сравнить аванс с обеспечительным по этому акту нельзя — нужен учёт обеспечительного.', true]);
    else if (excess > EPS) lines.push([`Сверх положенного обеспечительного переплачено ${fmt(excess)} — это и надо разбирать.`, true]);
    else if (excess < -EPS) lines.push([`Аванса не хватает до положенного обеспечительного на ${fmt(-excess)} — обеспечительный внесён не полностью или есть долг по аренде.`, true]);
    else lines.push(['Аванс ровно равен положенному обеспечительному — расчёты в порядке.', true]);
    return lines;
  }

  // ---------- весь разбор ----------
  // opts: { role, entity: 'ООО'|'ИП', half: 'auto'|'left'|'right', depositMonths, ourFile, theirFile, theirName }
  async function process(opts) {
    const R = ROLES[opts.role];
    const entRe = opts.entity === 'ИП' ? /(^|[\s(«"])ИП([\s)»",]|$)|индивидуальн/i : /(^|[\s«"])(ООО|АО|ПАО|ЗАО)([\s»",]|$)|обществ/i;
    const parse = async file => {
      const grid = await readGrid(file);
      const { hr, blocks } = findBlocks(grid);
      return { grid, hr, blocks: blocks.map(b => ({ b, label: blockLabel(grid, hr, b) })) };
    };
    const pickOur = blocks => {
      if (opts.half === 'left' || blocks.length === 1) return 0;
      if (opts.half === 'right') return blocks.length - 1;
      const hits = blocks.map((x, i) => entRe.test(x.label) ? i : -1).filter(i => i >= 0);
      return hits.length === 1 ? hits[0] : 0;
    };
    const tenantOur = opts.role === 'tenant';
    const ourP = await parse(opts.ourFile);
    const oi = pickOur(ourP.blocks);
    const ob = ourP.blocks[oi];
    const our = { name: ob.label || 'Наши данные', file: opts.ourFile.name, ...readBlock(ourP.grid, ourP.hr, ob.b, 'our', tenantOur) };
    if (!our.rows.length) throw new Error(`В файле «${opts.ourFile.name}» не нашлось операций с датой и суммой.`);

    let their = null, theirSource = '';
    if (opts.theirFile) {
      const tp = await parse(opts.theirFile);
      // в акте контрагента его данные — блок, который не похож на наше юрлицо (иначе левый)
      const ti = tp.blocks.length > 1 && entRe.test(tp.blocks[0].label) && !entRe.test(tp.blocks[1].label) ? 1 : 0;
      const tb = tp.blocks[opts.half === 'right' ? 0 : ti];
      their = { name: opts.theirName || tb.label || `Данные ${R.gen}`, file: opts.theirFile.name, ...readBlock(tp.grid, tp.hr, tb.b, 'their', !tenantOur) };
      theirSource = 'file';
    } else if (ourP.blocks.length > 1) {
      const tb = ourP.blocks[oi === 0 ? ourP.blocks.length - 1 : 0];
      const t = readBlock(ourP.grid, ourP.hr, tb.b, 'their', !tenantOur);
      if (t.rows.length) { their = { name: opts.theirName || tb.label || `Данные ${R.gen}`, file: opts.ourFile.name + ' (вторая половина)', ...t }; theirSource = 'half'; }
    }
    if (their && !their.rows.length) throw new Error(`В акте контрагента «${their.file}» не нашлось операций с датой и суммой.`);

    const cpName = opts.theirName || (their && their.name) || (ourP.blocks.length > 1 ? ourP.blocks[oi === 0 ? ourP.blocks.length - 1 : 0].label : '') || opts.ourFile.name.replace(/\.[^.]+$/, '');
    const groups = their ? match(our, their, R) : null;
    const months = analyzeMonths(our, 2, opts.depositMonths);
    const periodTxt = `${dstr(minDate([...our.rows, ...(their ? their.rows : [])]))} – ${dstr(maxDate([...our.rows, ...(their ? their.rows : [])]))}`;
    const warnings = [];
    [our, their].filter(Boolean).forEach(s => {
      const flipped = s.rows.filter(x => x.kind === PAY && /^(продажа|реализац)/i.test(x.doc)).length;
      if (flipped >= 2) warnings.push(`В акте «${s.name}» документы «Продажа/Реализация» попали в колонку оплат. Похоже, роль выбрана неверно — попробуйте «${opts.role === 'tenant' ? 'арендодатель' : 'арендатор'}».`);
      if (s.closingStated != null && Math.abs(s.closingStated - closingOf(s)) > 0.01)
        warnings.push(`В акте «${s.name}» сальдо на конец указано ${fmt(s.closingStated)}, а по операциям выходит ${fmt(closingOf(s))}. Часть строк не распознана — проверьте файл.`);
    });
    let mirror = false;
    if (their && groups.every(gr => ['ok'].includes(gr.status)) && theirSource === 'half') {
      mirror = true;
      warnings.push('Половины акта совпадают построчно. Если акт выгружен из нашей 1С, вторая половина заполнена автоматически и не подтверждает данные контрагента — для сверки нужен его подписанный акт.');
    }
    return { opts, R, our, their, theirSource, groups, months, warnings, mirror, cpName, periodTxt };
  }
  const minDate = rows => rows.reduce((a, r) => r.dt.n < a.n ? r.dt : a, rows[0].dt);
  const maxDate = rows => rows.reduce((a, r) => r.dt.n > a.n ? r.dt : a, rows[0].dt);

  // ---------- отчёт на странице ----------
  const MONTH_CLS = { ok: 'ok', shift: 'sh', deposit: 'dp', over: 'ov', under: 'un', nopay: 'np' };
  const GROUP_CLS = { ok: 'ok', shift: 'ok', split: 'sh', period: 'ov', round: 'ov', vat: 'ov', amount: 'un', their_acc: 'np', their_pay: 'np', our_acc: 'np', our_pay: 'np' };
  function renderResult(res) {
    const { R, our, their, groups, months, warnings } = res;
    const RS = reasons(R);
    const oc = closingOf(our);
    const h = [];
    h.push(`<div class="rc-head"><div><h2>${esc(res.cpName)}</h2><p class="help">Период ${esc(res.periodTxt)} · наше юрлицо: ${esc(res.opts.entity)} — ${res.opts.role === 'tenant' ? 'арендатор' : 'арендодатель'}</p></div>
      <button type="button" class="primary" id="rcDownload">Скачать разбор (Excel)</button></div>`);
    warnings.forEach(w => h.push(`<div class="notice bad">⚠ ${esc(w)}</div>`));
    // сальдо
    if (their) {
      const tc = closingOf(their);
      const diff = r2(tc - oc);
      h.push(`<div class="advice">${diff === 0 ? '<b>Сальдо сходится.</b>' : `<b>Сальдо не сходится на ${fmt(Math.abs(diff))} ₽.</b>`} ${R.debt}: у нас ${fmt(oc)}, у ${R.gen} ${fmt(tc)}${oc < 0 || tc < 0 ? ' (минус — аванс/переплата арендатора)' : ''}.</div>`);
    } else {
      h.push(`<div class="advice">${R.debt} на конец периода: <b class="num">${fmt(oc)} ₽</b>${oc < 0 ? ' (минус — аванс/переплата арендатора)' : ''}. Акта контрагента нет — показан разбор наших данных по месяцам.</div>`);
    }
    // расхождения сторон
    if (their) {
      const bad = groups.filter(gr => !['ok', 'shift', 'split'].includes(gr.status));
      h.push(`<h3>Расхождения между сторонами</h3>`);
      if (!bad.length) h.push(`<p class="help">Все операции совпали${groups.some(g => g.status === 'split') ? ' (некоторые — одна оплата у одной стороны = несколько у другой)' : ''}.</p>`);
      else h.push(table(['Что', 'Дата', 'Документ', 'Сумма', 'Влияние', 'Что сделать'], bad.sort((a, b) => Math.abs(b.impact) - Math.abs(a.impact)).map(gr => {
        const x = (gr.their[0] || gr.our[0]);
        return { cls: GROUP_CLS[gr.status], cells: [RS[gr.status][0], dstr(x.dt), x.doc + (gr.note ? ` — ${gr.note}` : ''), fmt(x.amt), fmt(gr.impact), RS[gr.status][1]] };
      }), [3, 4]));
    }
    // по месяцам
    const { periods, dep } = months;
    h.push(`<h3>Месяцы без оплаты, переплаты и недоплаты</h3>`);
    if (!periods.length) h.push('<p class="help">Не найдено: каждый месяц оплачен на сумму начисления (с учётом оплат вперёд и обеспечительного).</p>');
    else {
      h.push(table(['Период', 'Что', 'Сумма', 'Подробно', 'Что сделать'], periods.map(p => ({ cls: MONTH_CLS[p.status], cells: [p.label, MONTH_STATUS[p.status], fmt(Math.abs(p.total)), p.detail, R[p.status]] })), [2]));
      const over = periods.filter(p => p.total > 0).reduce((a, p) => a + p.total, 0), under = -periods.filter(p => p.total < 0).reduce((a, p) => a + p.total, 0);
      h.push(`<p class="help">Итого переплачено ${fmt(over)}, недоплачено ${fmt(under)}.</p>`);
    }
    // обеспечительный
    if (dep.finalRate) {
      h.push(`<h3>Обеспечительный платёж</h3>`);
      if (dep.events.length) h.push(table(['Событие', 'Месяц', 'Ожидалось', 'Найдено', 'Не хватает'], dep.events.map(e => {
        const short = r2(Math.max(0, e.expected - e.found));
        return { cls: short > EPS ? 'un' : 'dp', cells: [e.what, e.when, fmt(e.expected), fmt(e.found), fmt(short)] };
      }), [2, 3, 4]));
      depositLines(our, dep, periods).forEach(([t, b]) => h.push(`<p class="${b ? 'rc-strong' : 'help'}">${esc(t)}</p>`));
    }
    h.push(`<details class="set"><summary>Все месяцы <span>(${months.ms.length})</span></summary>${table(['Месяц', 'Ставка', 'Начислено', 'Оплачено', 'Расхождение', 'Статус', 'Комментарий'],
      months.ms.map(mo => ({ cls: MONTH_CLS[mo.status] || '', cells: [mlabel(mo.k), regularRate(mo) == null ? '' : fmt(regularRate(mo)), fmt(mo.acc), fmt(mo.pay), fmt(mo.res), MONTH_STATUS[mo.status], mo.notes.join('; ')] })), [1, 2, 3, 4])}</details>`);
    h.push(`<p class="help">Распределение по месяцам и обеспечительному сделано по суммам и датам; назначение платежей программа не видит. Проверяйте выводы по договору и выписке.</p>`);
    return h.join('');
  }
  function table(head, rows, numCols = []) {
    return `<div class="rc-scroll"><table class="rc-table"><thead><tr>${head.map(x => `<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>${rows.map(r =>
      `<tr class="${r.cls || ''}">${r.cells.map((c, i) => `<td class="${numCols.includes(i) ? 'num' : ''}">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  }

  // ---------- Excel ----------
  const FILL = { ok: 'FFE2F0D9', sh: 'FFDDEBF7', dp: 'FFE4DFEC', ov: 'FFFFF2CC', un: 'FFFCE4D6', np: 'FFF8D7DA', head: 'FFD9D9D9' };
  const MONEY = '#,##0.00;[Red]-#,##0.00';
  async function buildXlsx(res) {
    await loadScript(LIB.exceljs);
    const { R, our, their, groups, months } = res;
    const RS = reasons(R);
    const wb = new window.ExcelJS.Workbook();
    const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
    const put = (ws, r, c, v, o = {}) => {
      const x = ws.getCell(r, c); x.value = v;
      if (o.border !== false) x.border = { top: thin, left: thin, bottom: thin, right: thin };
      if (o.fill) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: FILL[o.fill] || o.fill } };
      if (o.money) x.numFmt = MONEY;
      if (o.date) x.numFmt = 'dd.mm.yyyy';
      if (o.wrap) x.alignment = { wrapText: true, vertical: 'top' };
      if (o.bold) x.font = { bold: true };
      return x;
    };
    const head = (ws, r, titles) => titles.forEach((t, i) => put(ws, r, i + 1, t, { fill: 'head', bold: true, wrap: true }));
    const plain = (ws, r, t, font = {}) => { const x = ws.getCell(r, 1); x.value = t; x.font = font; };

    // Итог
    const ws = wb.addWorksheet('Итог');
    plain(ws, 1, their ? `Сверка расчётов: ${res.cpName}` : `Разбор оплат по месяцам: ${res.cpName}`, { bold: true, size: 14 });
    plain(ws, 2, `Период ${res.periodTxt} · наше юрлицо: ${res.opts.entity} — ${res.opts.role === 'tenant' ? 'арендатор' : 'арендодатель'}`);
    plain(ws, 3, `Файлы: ${our.file}${their ? '; ' + their.file : ''}`, { italic: true, color: { argb: 'FF808080' } });
    let r = 5;
    const sum = (s, k) => s.rows.filter(x => x.kind === k).reduce((a, x) => a + x.amt, 0);
    head(ws, r, their ? ['Показатель', 'У нас', `У ${R.gen}`, 'Разница'] : ['Показатель', 'Сумма']);
    [[`${R.debt} на начало периода`, our.opening, their && their.opening], ['Начислено', sum(our, ACC), their && sum(their, ACC)],
     ['Оплачено', sum(our, PAY), their && sum(their, PAY)], [`${R.debt} на конец периода`, closingOf(our), their && closingOf(their)]].forEach(([l, a, b], i, arr) => {
      r++; put(ws, r, 1, l, { bold: i === arr.length - 1 }); put(ws, r, 2, r2(a), { money: true, bold: i === arr.length - 1 });
      if (their) { put(ws, r, 3, r2(b), { money: true }); put(ws, r, 4, r2(b - a), { money: true }); }
    });
    r++; plain(ws, r, 'Отрицательный долг = переплата (аванс) арендатора.', { italic: true, color: { argb: 'FF808080' } });
    res.warnings.forEach(w => { r++; plain(ws, r, '⚠ ' + w, { color: { argb: 'FFC00000' } }); });
    if (their) {
      r += 2; plain(ws, r, 'Из чего складывается разница между сторонами', { bold: true, size: 12 });
      r++; head(ws, r, ['Причина', 'Кол-во', 'Сумма влияния', 'Что сделать']);
      const openDiff = r2(their.opening - our.opening);
      if (Math.abs(openDiff) > 0.01) { r++; ['Разное сальдо на начало периода', 1, openDiff, 'Сначала сверить предыдущий период.'].forEach((v, i) => put(ws, r, i + 1, v, { wrap: true, money: i === 2 })); }
      const by = new Map();
      groups.filter(gr => !['ok', 'shift', 'split'].includes(gr.status)).forEach(gr => { const s = by.get(gr.status) || [0, 0]; s[0]++; s[1] += gr.impact; by.set(gr.status, s); });
      by.forEach(([cnt, imp], st) => { r++; [RS[st][0], cnt, r2(imp), RS[st][1]].forEach((v, i) => put(ws, r, i + 1, v, { wrap: true, money: i === 2, fill: GROUP_CLS[st] })); });
      const explained = r2(openDiff + groups.reduce((a, gr) => a + gr.impact, 0));
      r++; plain(ws, r, 'Не объяснено', { bold: true }); put(ws, r, 3, r2(closingOf(their) - closingOf(our) - explained), { money: true, bold: true, border: false });
    }
    r += 2; plain(ws, r, `Месяцы без оплаты, переплаты и недоплаты — ${our.name}`, { bold: true, size: 12 });
    const { periods, dep } = months;
    if (!periods.length) { r++; plain(ws, r, 'Не найдено: каждый месяц оплачен на сумму начисления (с учётом оплат вперёд и обеспечительного).'); }
    else {
      r++; head(ws, r, ['Период', 'Что', 'Сумма', 'Подробно', 'Что сделать']);
      periods.forEach(p => { r++; [p.label, MONTH_STATUS[p.status], Math.abs(p.total), p.detail, R[p.status]].forEach((v, i) => put(ws, r, i + 1, v, { wrap: true, money: i === 2, fill: MONTH_CLS[p.status] })); });
    }
    if (dep.finalRate) {
      r += 2; plain(ws, r, 'Обеспечительный платёж', { bold: true, size: 12 });
      if (dep.events.length) {
        r++; head(ws, r, ['Событие', 'Месяц', 'Ожидалось', 'Найдено в оплатах', 'Не хватает']);
        dep.events.forEach(e => { r++; const short = r2(Math.max(0, e.expected - e.found)); [e.what, e.when, e.expected, e.found, short].forEach((v, i) => put(ws, r, i + 1, v, { wrap: true, money: i >= 2, fill: short > EPS ? 'un' : 'dp' })); });
      }
      depositLines(our, dep, periods).forEach(([t, b]) => { r++; plain(ws, r, t, b ? { bold: true } : {}); });
    }
    [40, 22, 18, 50, 70].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    // Сопоставление и Что сделать
    if (their) {
      const w2 = wb.addWorksheet('Сопоставление');
      plain(w2, 1, our.name, { bold: true }); w2.getCell(1, 6).value = their.name; w2.getCell(1, 6).font = { bold: true };
      head(w2, 2, ['Дата', 'Документ', 'Начислено', 'Оплачено', 'Статус', 'Дата', 'Документ', 'Начислено', 'Оплачено', 'Комментарий']);
      w2.views = [{ state: 'frozen', ySplit: 2 }];
      let rr = 2;
      groups.forEach(gr => {
        const n = Math.max(gr.our.length, gr.their.length);
        for (let k = 0; k < n; k++) {
          rr++;
          const f = GROUP_CLS[gr.status];
          for (let c = 1; c <= 10; c++) put(w2, rr, c, null, { fill: f });
          [[gr.our[k], 1], [gr.their[k], 6]].forEach(([x, base]) => {
            if (!x) return;
            put(w2, rr, base, xlDate(x.dt), { date: true, fill: f }); put(w2, rr, base + 1, x.doc, { fill: f });
            put(w2, rr, base + (x.kind === ACC ? 2 : 3), x.amt, { money: true, fill: f });
          });
          put(w2, rr, 5, k === 0 ? RS[gr.status][0] : '↳', { fill: f }); put(w2, rr, 10, k === 0 ? gr.note : '', { fill: f, wrap: true });
        }
      });
      [11, 40, 14, 14, 24, 11, 40, 14, 14, 50].forEach((w, i) => { w2.getColumn(i + 1).width = w; });
      w2.autoFilter = `A2:J${rr}`;

      const w3 = wb.addWorksheet('Что сделать');
      head(w3, 1, ['Причина', 'Дата', 'Документ', 'Сторона', 'Сумма', 'Влияние на разницу', 'Подсказка', 'Что сделать']);
      let r3 = 1;
      groups.filter(gr => !['ok', 'shift', 'split'].includes(gr.status)).forEach(gr => {
        const x = gr.their[0] || gr.our[0]; r3++;
        const sideTxt = gr.their.length && !gr.our.length ? R.cp : !gr.their.length ? 'мы' : 'обе';
        [RS[gr.status][0], xlDate(x.dt), x.doc, sideTxt, x.amt, gr.impact, gr.note, RS[gr.status][1]]
          .forEach((v, i) => put(w3, r3, i + 1, v, { wrap: true, fill: GROUP_CLS[gr.status], money: i === 4 || i === 5, date: i === 1 }));
      });
      [26, 11, 40, 13, 14, 14, 40, 60].forEach((w, i) => { w3.getColumn(i + 1).width = w; });
    }

    // По месяцам
    const w4 = wb.addWorksheet('По месяцам');
    plain(w4, 1, our.name, { bold: true });
    head(w4, 2, ['Месяц', 'Ставка', 'Начислено', 'Оплачено', 'Разница за месяц', 'Сдвиг из/в соседние месяцы', 'Обеспечительный', 'Расхождение', `${R.debt} на конец месяца`, 'Статус', 'Комментарий']);
    w4.views = [{ state: 'frozen', ySplit: 2 }];
    let r4 = 2;
    months.ms.forEach(mo => {
      r4++; const f = MONTH_CLS[mo.status];
      [mlabel(mo.k), regularRate(mo), mo.acc, mo.pay, mo.diff, r2(mo.res + mo.dep - mo.diff), mo.dep, mo.res, mo.balance, MONTH_STATUS[mo.status], mo.notes.join('; ')]
        .forEach((v, i) => put(w4, r4, i + 1, v, { fill: f, money: i >= 1 && i <= 8, wrap: i === 10 }));
    });
    [11, 13, 13, 13, 14, 15, 15, 13, 18, 26, 60].forEach((w, i) => { w4.getColumn(i + 1).width = w; });
    w4.autoFilter = `A2:K${r4}`;

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  const fileName = res => {
    const t = new Date(), d = `${String(t.getDate()).padStart(2, '0')}.${String(t.getMonth() + 1).padStart(2, '0')}.${t.getFullYear()}`;
    const who = String(res.cpName || 'контрагент').replace(/[\\/:*?"<>|«»]+/g, '').replace(/\s+/g, '_').slice(0, 60);
    return `Разбор_${who}_${d}.xlsx`;
  };

  // ---------- форма на вкладке ----------
  const PREF_KEY = 'arenda-sverka-prefs';
  let mounted = false;
  function mount(root) {
    if (mounted) return;
    mounted = true;
    let prefs = {};
    try { prefs = JSON.parse(localStorage.getItem(PREF_KEY) || '{}') || {}; } catch (e) {}
    root.innerHTML = `
      <form class="f card" id="rcForm" novalidate>
        <h2>Сверка акта</h2>
        <p class="help">Загрузите акт сверки — программа найдёт месяцы без оплаты, переплаты и недоплаты, учтёт обеспечительный платёж и, если есть акт контрагента, сравнит стороны. Файлы обрабатываются в браузере и никуда не отправляются.</p>
        <div class="grid">
          <fieldset><legend class="help">Наше юрлицо в акте</legend>
            <label class="check"><input type="radio" name="rcEnt" value="ООО"> ООО</label>
            <label class="check"><input type="radio" name="rcEnt" value="ИП"> ИП</label></fieldset>
          <fieldset><legend class="help">Наше юрлицо в договоре</legend>
            <label class="check"><input type="radio" name="rcRole" value="tenant"> Арендатор — мы платим аренду</label>
            <label class="check"><input type="radio" name="rcRole" value="landlord"> Арендодатель — нам платят аренду</label></fieldset>
        </div>
        <label>Наш акт сверки (Excel, CSV или PDF из 1С)<input type="file" id="rcOur" accept=".xlsx,.xls,.xlsm,.csv,.pdf"></label>
        <label>Акт контрагента — если есть (Excel, CSV или PDF)<input type="file" id="rcTheir" accept=".xlsx,.xls,.xlsm,.csv,.pdf"></label>
        <details class="set"><summary>Дополнительно <span>· обеспечительный, половина акта</span></summary>
          <div class="grid">
            <label>Обеспечительный, месячных ставок<input type="number" id="rcDep" min="0" max="12" step="0.5" value="1"></label>
            <label>Наши данные в акте<select id="rcHalf"><option value="auto">Найти по юрлицу</option><option value="left">Левая половина</option><option value="right">Правая половина</option></select></label>
            <label class="wide">Название контрагента для отчёта<input id="rcName" placeholder="Возьмём из акта" autocomplete="off"></label>
          </div>
          <p class="help">Обеспечительный обычно равен месячной аренде и доплачивается при индексации. Укажите 0, если по договору его нет.</p>
        </details>
        <div class="err" id="rcErr" hidden></div>
        <div class="f-actions"><button type="submit" class="primary" id="rcGo">Сделать разбор</button></div>
      </form>
      <div id="rcOut" class="card rc-out" hidden></div>`;
    const q = s => root.querySelector(s);
    const setRadio = (name, v) => { const el = root.querySelector(`input[name="${name}"][value="${v}"]`); if (el) el.checked = true; };
    setRadio('rcEnt', prefs.entity || 'ООО'); setRadio('rcRole', prefs.role || 'tenant');
    if (prefs.dep != null) q('#rcDep').value = prefs.dep;
    let last = null;
    q('#rcForm').addEventListener('submit', async e => {
      e.preventDefault();
      const err = m => { q('#rcErr').textContent = m; q('#rcErr').hidden = false; };
      q('#rcErr').hidden = true;
      const ourFile = q('#rcOur').files[0], theirFile = q('#rcTheir').files[0];
      if (!ourFile) return err('Выберите файл нашего акта сверки.');
      const opts = {
        entity: root.querySelector('input[name="rcEnt"]:checked').value,
        role: root.querySelector('input[name="rcRole"]:checked').value,
        depositMonths: Math.max(0, +q('#rcDep').value || 0), half: q('#rcHalf').value,
        theirName: q('#rcName').value.trim(), ourFile, theirFile,
      };
      try { localStorage.setItem(PREF_KEY, JSON.stringify({ entity: opts.entity, role: opts.role, dep: opts.depositMonths })); } catch (x) {}
      const btn = q('#rcGo'); btn.disabled = true; btn.textContent = 'Считаю…';
      try {
        last = await process(opts);
        const out = q('#rcOut'); out.innerHTML = renderResult(last); out.hidden = false;
        out.scrollIntoView({ behavior: 'smooth', block: 'start' });
      } catch (x) { err(x.message || 'Не удалось разобрать файл.'); q('#rcOut').hidden = true; }
      finally { btn.disabled = false; btn.textContent = 'Сделать разбор'; }
    });
    q('#rcOut').addEventListener('click', async e => {
      if (!e.target.closest('#rcDownload') || !last) return;
      const b = e.target.closest('#rcDownload'); b.disabled = true; b.textContent = 'Готовлю файл…';
      try {
        const blob = await buildXlsx(last);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = fileName(last);
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 4000);
      } catch (x) { const n = document.createElement('div'); n.className = 'notice bad'; n.textContent = x.message || 'Не удалось сформировать файл.'; b.after(n); }
      finally { b.disabled = false; b.textContent = 'Скачать разбор (Excel)'; }
    });
  }

  window.Sverka = { mount, process, buildXlsx, renderResult, readGrid, fileName };
})();
