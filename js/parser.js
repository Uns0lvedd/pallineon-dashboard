/*
 * Parser for the property manager's monthly owner statement (xlsx).
 * One shared module: the GitHub Action uses it to build data.json, and the
 * dashboard uses it in the browser for the drag-and-drop preview.
 *
 * It locates the header row and columns by their titles, so the layout may
 * shift between months. It tolerates numbers stored as text ("2,357.00"),
 * formulas, "-" for empty, and the internet charge typed as a date (28.9).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReportParser = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var MONTH_NAMES = ['january', 'february', 'march', 'april', 'may', 'june', 'july',
    'august', 'september', 'october', 'november', 'december'];
  var READ_OPTIONS = { cellNF: true, cellText: true };

  // Money fields kept for every platform row, in statement order.
  var FIELDS = ['gross', 'commission', 'netOfCommission', 'nights', 'envFee', 'afterFee',
    'vat', 'cityTax', 'mgmtFee', 'orders', 'cleaning', 'internet', 'improvements', 'ownerPayment'];

  var COLUMN_PATTERNS = [
    ['source', /reservation\s*source/],
    ['month', /^month$/],
    ['commission', /commission/],
    ['nights', /nig?h?th?|night/],
    ['envFee', /environment/],
    ['afterFee', /income\s*after/],
    ['vat', /^vat$/],
    ['cityTax', /city\s*tax/],
    ['mgmtFee', /management/],
    ['orders', /orders/],
    ['cleaning', /cleaning/],
    ['internet', /internet/],
    ['improvements', /improvement|fault/],
    ['ownerPayment', /owner\s*payment/]
  ];
  var REQUIRED = ['gross', 'commission', 'netOfCommission', 'nights', 'envFee', 'afterFee', 'vat',
    'cityTax', 'mgmtFee', 'orders', 'cleaning', 'internet', 'improvements', 'ownerPayment'];
  var COLUMN_TITLES = {
    gross: 'INCOME GROSS EURO (first)', commission: 'commission', netOfCommission: 'INCOME GROSS EURO (second)',
    nights: 'Total night per month', envFee: 'Environmental Fee', afterFee: 'INCOME after fee', vat: 'VAT',
    cityTax: 'City TAX', mgmtFee: 'Management Fee', orders: 'Number of Orders', cleaning: 'Cleaning cost',
    internet: 'Internet Charge', improvements: 'Improvements / Faults', ownerPayment: 'Owner Payment'
  };

  function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
  function norm(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase(); }

  function monthFromFilename(name) {
    var base = String(name).replace(/\.[a-z0-9]+$/i, '').toLowerCase();
    var re = /(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept|sep|oct|nov|dec)[^a-z0-9]*(20\d{2})/g;
    var m, last = null;
    while ((m = re.exec(base))) last = m;
    if (last) {
      var key = last[1].slice(0, 3);
      for (var i = 0; i < 12; i++) if (MONTH_NAMES[i].slice(0, 3) === key) return { year: +last[2], month: i + 1 };
    }
    var iso = /(20\d{2})[-_.](0[1-9]|1[0-2])(?!\d)/.exec(base);
    if (iso) return { year: +iso[1], month: +iso[2] };
    return null;
  }

  function monthKey(year, month) { return year + '-' + (month < 10 ? '0' : '') + month; }

  function parseNumberText(s) {
    s = String(s).trim();
    if (s === '' || s === '-' || s === '–') return null;
    var pct = /%$/.test(s);
    s = s.replace(/[%€₪\s]/g, '');
    var hasComma = s.indexOf(',') >= 0, hasDot = s.indexOf('.') >= 0;
    if (hasComma && hasDot) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.'); // 1.296,00
      else s = s.replace(/,/g, '');                                                             // 2,357.00
    } else if (hasComma) {
      if (/^-?\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, ''); else s = s.replace(',', '.');
    }
    if (!/^-?\d+(\.\d+)?$/.test(s)) return undefined; // not a number
    var n = parseFloat(s);
    return pct ? n / 100 : n;
  }

  function parseWorkbook(XLSX, workbook, filename) {
    var warnings = [];
    var fm = monthFromFilename(filename);
    if (!fm) throw new Error('The file name must contain the month and year, for example "PALLINEON_R_1_AUGUST_2026.xlsx".');

    var ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws || !ws['!ref']) throw new Error('The first sheet is empty.');
    var range = XLSX.utils.decode_range(ws['!ref']);
    var maxRow = Math.min(range.e.r, 400), maxCol = Math.min(range.e.c, 40);

    function cell(r, c) { return ws[XLSX.utils.encode_cell({ r: r, c: c })]; }
    function text(r, c) { var x = cell(r, c); return x && x.t === 's' ? String(x.v).trim() : (x && x.w ? String(x.w).trim() : ''); }
    function isDateCell(x) { return !!(x && x.t === 'n' && x.z && XLSX.SSF.is_date(x.z)); }
    // Returns a number, null for empty, or {date:{m,d}} when a number was typed as a date.
    function value(r, c) {
      var x = cell(r, c);
      if (!x || x.v == null) return null;
      if (x.t === 'n') {
        if (isDateCell(x)) { var d = XLSX.SSF.parse_date_code(x.v); return { date: { m: d.m, d: d.d } }; }
        return x.v;
      }
      if (x.t === 's') { var n = parseNumberText(x.v); return n === undefined ? null : n; }
      return null;
    }
    function plain(v) { return typeof v === 'number' ? v : null; }

    // ---- locate header rows -------------------------------------------------
    var headers = [];
    for (var r = range.s.r; r <= maxRow; r++) {
      var cols = {}, grossSeen = 0, hits = 0;
      for (var c = range.s.c; c <= maxCol; c++) {
        var t = norm(text(r, c));
        if (!t) continue;
        if (/income\s*gross/.test(t)) { cols[grossSeen === 0 ? 'gross' : 'netOfCommission'] = c; grossSeen++; hits++; continue; }
        for (var i = 0; i < COLUMN_PATTERNS.length; i++) {
          var key = COLUMN_PATTERNS[i][0];
          if (cols[key] == null && COLUMN_PATTERNS[i][1].test(t)) { cols[key] = c; hits++; break; }
        }
      }
      if (hits >= 6 && cols.ownerPayment != null) headers.push({ row: r, cols: cols });
    }
    if (!headers.length) throw new Error('Could not find the header row (the one with "Reservation Source", "Month", "Owner Payment").');

    headers.forEach(function (h) {
      var missing = REQUIRED.filter(function (k) { return h.cols[k] == null; });
      if (missing.length) {
        throw new Error('Missing or renamed column(s) in row ' + (h.row + 1) + ': ' +
          missing.map(function (k) { return COLUMN_TITLES[k]; }).join(', ') + '.');
      }
      if (h.cols.month == null) h.cols.month = (h.cols.source != null ? h.cols.source : 0) + 1;
      if (h.cols.source == null) h.cols.source = h.cols.month - 1;
    });

    // ---- rates (the row above the first header) -----------------------------
    var h0 = headers[0];
    var rates = {
      envFeePerNight: plain(value(h0.row - 1, h0.cols.envFee)),
      vat: plain(value(h0.row - 1, h0.cols.vat)),
      cityTax: plain(value(h0.row - 1, h0.cols.cityTax)),
      mgmtFee: plain(value(h0.row - 1, h0.cols.mgmtFee)),
      cleaningPerOrder: plain(value(h0.row - 1, h0.cols.cleaning)),
      commission: {}
    };

    // ---- blocks: one per platform -------------------------------------------
    var platforms = {}, otherMonthsWithData = [];
    headers.forEach(function (h, hi) {
      var end = hi + 1 < headers.length ? headers[hi + 1].row : maxRow + 1;
      var rows = [], totalRow = null, label = '';
      for (var r = h.row + 1; r < end; r++) {
        var m = norm(text(r, h.cols.month)), s = text(r, h.cols.source);
        if (m === 'total') { totalRow = r; break; }
        if (!s && !m) { if (rows.length) break; else continue; }
        if (!label && s) label = s;
        rows.push(r);
      }
      if (!rows.length) return;
      var key = norm(label).replace(/[^a-z0-9]+/g, '') || ('platform' + (hi + 1));
      var commissionRate = plain(value(h.row - 1, h.cols.commission));
      if (commissionRate != null) rates.commission[key] = commissionRate;

      function readRow(r) {
        var out = {}, any = false, repaired = [];
        FIELDS.forEach(function (f) {
          var v = value(r, h.cols[f]);
          if (v && v.date) {
            // A decimal such as 28.9 that Excel turned into the date 28 September.
            var fromTotal = totalRow != null ? plain(value(totalRow, h.cols[f])) : null;
            v = fromTotal != null ? fromTotal : parseFloat(v.date.d + '.' + v.date.m);
            repaired.push(f);
          }
          out[f] = v == null ? 0 : round2(v);
          if (out[f] !== 0) any = true;
        });
        return { data: out, any: any, repaired: repaired };
      }

      var chosen = null;
      rows.forEach(function (r, idx) {
        var monthNo = idx + 1; // rows run January to December
        var row = readRow(r);
        if (!row.any) return;
        if (monthNo === fm.month) { chosen = row; chosen.r = r; }
        else if (otherMonthsWithData.indexOf(monthNo) < 0) otherMonthsWithData.push(monthNo);
      });

      var data = chosen ? chosen.data : FIELDS.reduce(function (o, f) { o[f] = 0; return o; }, {});
      data.label = label;
      if (chosen) {
        var noteCol = h.cols.ownerPayment + 1, note = [];
        for (var c = noteCol; c <= Math.min(noteCol + 3, maxCol); c++) {
          var x = cell(chosen.r, c);
          if (x && x.t === 's' && String(x.v).trim()) note.push(String(x.v).trim());
        }
        if (note.length) data.note = note.join(' ');
        if (chosen.repaired.length) {
          warnings.push(label + ': ' + chosen.repaired.map(function (f) { return COLUMN_TITLES[f]; }).join(', ') +
            ' was stored as a date in the sheet and was read as ' + chosen.repaired.map(function (f) { return data[f]; }).join(', ') + '.');
        }
      }
      platforms[key] = data;
    });

    if (!Object.keys(platforms).length) throw new Error('No platform rows found under the header.');
    var hasData = Object.keys(platforms).some(function (k) {
      return FIELDS.some(function (f) { return platforms[k][f] !== 0; });
    });
    if (!hasData && otherMonthsWithData.length) {
      throw new Error('The file name says ' + MONTH_NAMES[fm.month - 1] + ' ' + fm.year + ', but the data is in the ' +
        otherMonthsWithData.map(function (m) { return MONTH_NAMES[m - 1]; }).join(', ') + ' row. Rename the file or ask the manager for a corrected one.');
    }
    if (otherMonthsWithData.length) {
      warnings.push('Rows for other months (' + otherMonthsWithData.map(function (m) { return MONTH_NAMES[m - 1]; }).join(', ') +
        ') also contain data and were ignored. Each file is expected to hold one month.');
    }
    if (!hasData) warnings.push('No figures found for ' + MONTH_NAMES[fm.month - 1] + '. The month is recorded with zeros.');

    // ---- "TOTAL all Platforms" ----------------------------------------------
    var statedTotal = null;
    for (var rr = headers[headers.length - 1].row; rr <= maxRow && statedTotal == null; rr++) {
      for (var cc = range.s.c; cc <= maxCol; cc++) {
        if (/total all platforms/.test(norm(text(rr, cc)))) { statedTotal = plain(value(rr, h0.cols.ownerPayment)); break; }
      }
    }

    return finalize({
      month: monthKey(fm.year, fm.month), source: String(filename), rates: rates,
      platforms: platforms, statedTotal: statedTotal == null ? null : round2(statedTotal), warnings: warnings
    });
  }

  /* Recompute every derived line from the rates and record any that differ
     from what the statement says. Used for parsed and hand-entered months. */
  function finalize(m) {
    var r = m.rates || {}, checks = [], TOL = 0.015, sum = 0;
    Object.keys(m.platforms).forEach(function (key) {
      var p = m.platforms[key];
      FIELDS.forEach(function (f) { p[f] = round2(+p[f] || 0); });
      sum += p.ownerPayment;
      function check(field, expected) {
        if (expected == null || isNaN(expected)) return;
        expected = round2(expected);
        if (Math.abs(expected - p[field]) > TOL) checks.push({ platform: key, field: field, stated: p[field], expected: expected });
      }
      var cr = r.commission && r.commission[key];
      if (cr != null) check('commission', p.gross * cr);
      check('netOfCommission', p.gross - p.commission);
      if (r.envFeePerNight != null) check('envFee', p.nights * r.envFeePerNight);
      check('afterFee', p.netOfCommission - p.envFee);
      if (r.vat != null) check('vat', p.afterFee * r.vat);
      if (r.cityTax != null) check('cityTax', p.afterFee * r.cityTax);
      if (r.cleaningPerOrder != null) check('cleaning', p.orders * r.cleaningPerOrder);
      if (r.mgmtFee != null) check('mgmtFee', (p.netOfCommission - p.cleaning) * r.mgmtFee);
      check('ownerPayment', p.afterFee - p.vat - p.cityTax - p.mgmtFee - p.cleaning - p.internet - p.improvements);
    });
    m.ownerPayment = round2(sum);
    if (m.statedTotal != null && Math.abs(m.statedTotal - m.ownerPayment) > TOL) {
      checks.push({ platform: 'all', field: 'total', stated: m.statedTotal, expected: m.ownerPayment });
    }
    m.checks = checks;
    m.warnings = m.warnings || [];
    return m;
  }

  return {
    READ_OPTIONS: READ_OPTIONS, FIELDS: FIELDS, MONTH_NAMES: MONTH_NAMES,
    monthFromFilename: monthFromFilename, parseWorkbook: parseWorkbook, finalize: finalize, round2: round2
  };
});
