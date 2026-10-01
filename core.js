/*
 * SCAF -> Bộ chứng từ xuất khẩu (CI / PLIST / INV-3RD / INV - delta / Custom code)
 * Chạy được cả trên trình duyệt (window.ScafCore) và Node (module.exports).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ScafCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------- helpers
  function cellText(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
      if ('result' in v) return cellText(v.result);
      if (v.text !== undefined) return String(v.text);
      if (v.formula) return '';
      if (v.error) return '';
    }
    return String(v);
  }
  function cellRaw(v) {
    if (v && typeof v === 'object' && !(v instanceof Date)) {
      if (v.richText) return cellText(v);
      if ('result' in v) return v.result;
      if (v.text !== undefined) return v.text;
    }
    return v;
  }
  function norm(s) { return cellText(s).toLowerCase().replace(/[\s\.\/:_-]+/g, ' ').trim(); }
  function num(v) {
    v = cellRaw(v);
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }
  var MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  function parseDate(v) {
    v = cellRaw(v);
    if (v instanceof Date) return v;
    if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400000));
    var s = cellText(v).trim();
    var m = s.match(/(\d{1,2})[-\s\/]([A-Za-z]{3})[A-Za-z]*[-\s\/](\d{2,4})/);
    if (m && MONTHS[m[2].toLowerCase()] !== undefined) {
      var y = +m[3]; if (y < 100) y += 2000;
      return new Date(Date.UTC(y, MONTHS[m[2].toLowerCase()], +m[1]));
    }
    m = s.match(/(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{4})/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return null;
  }
  function fmtDate(d) {
    if (!d) return '';
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }

  /** "90(S)" -> mode 'letter' => "S", 'number' => 90, 'raw' => "90(S)" */
  function convertSize(size, mode) {
    var s = cellText(size).trim();
    var m = s.match(/^([^()]+?)\s*\(\s*([^()]+?)\s*\)$/);
    var out = s;
    if (m && mode !== 'raw') {
      var a = m[1].trim(), b = m[2].trim();
      var aNum = /^\d+(\.\d+)?$/.test(a), bNum = /^\d+(\.\d+)?$/.test(b);
      if (mode === 'number') out = aNum ? a : (bNum ? b : a);
      else out = !aNum ? a : (!bNum ? b : a); // letter
    }
    if (/^\d+(\.\d+)?$/.test(out)) return Number(out);
    return out;
  }

  // ----------------------------------------------------------- parse SCAF
  var COLS = {
    orderNo: ['order no'], story: ['story'], reference: ['reference'], kind: ['kind'],
    color: ['color', 'colour'], size: ['size'], ctnFrom: ['ctn from'], ctnTo: ['ctn to'],
    qty: ['qty'], boxes: ['number of box'], realQty: ['real qty'], boxDim: ['small box size']
  };

  /** ws: ExcelJS worksheet of the SCAF "PACKING LIST" export */
  function parseScafSheet(ws) {
    var grid = [];
    ws.eachRow({ includeEmpty: true }, function (row, r) {
      var arr = [];
      row.eachCell({ includeEmpty: true }, function (cell, c) { arr[c] = cell.value; });
      grid[r] = arr;
    });
    var maxR = grid.length;
    function nextValueRight(r, c, span) {
      var row = grid[r] || [];
      var end = Math.min(row.length, c + 1 + (span || 99));
      for (var k = c + 1; k < end; k++) if (cellText(row[k]).trim() !== '') return row[k];
      return null;
    }
    var info = { invoiceNo: '', invoiceDate: null, sailingDate: null, from: '', to: '', buyer1: '', buyer2: '', shipperPer: '' };
    var headerRow = -1, colMap = {};
    for (var r = 1; r < maxR; r++) {
      var row = grid[r] || [];
      for (var c = 1; c < row.length; c++) {
        var t = norm(row[c]);
        if (!t) continue;
        if (t.indexOf('inv no') === 0 || t === 'inv no date' || t.indexOf('invoice no') === 0) {
          var v = cellText(nextValueRight(r, c));
          var m = v.match(/([A-Z]{0,3}\d{5,})/i); if (m) info.invoiceNo = m[1];
          var dm = v.match(/date\s*:?\s*([^)]+)/i); if (dm) info.invoiceDate = parseDate(dm[1]);
        } else if (t.indexOf('sailling on') === 0 || t.indexOf('sailing on') === 0) {
          info.sailingDate = parseDate(nextValueRight(r, c));
        } else if (t === 'from') { info.from = cellText(nextValueRight(r, c, 8)).trim(); }
        else if (t === 'to') { info.to = cellText(nextValueRight(r, c, 8)).trim(); }
        else if (t.indexOf('shipper per') === 0) { info.shipperPer = cellText(nextValueRight(r, c, 8)).trim(); }
        else if (t.indexOf('messers') === 0) {
          info.buyer1 = cellText((grid[r + 1] || [])[c]).trim();
          info.buyer2 = cellText((grid[r + 2] || [])[c]).trim();
        } else if (t === 'order no' && headerRow < 0) {
          headerRow = r;
        }
      }
      if (headerRow === r) {
        for (var c2 = 1; c2 < row.length; c2++) {
          var h = norm(row[c2]);
          Object.keys(COLS).forEach(function (k) {
            if (colMap[k] === undefined && COLS[k].indexOf(h) >= 0) colMap[k] = c2;
          });
        }
      }
    }
    if (headerRow < 0) throw new Error('Không tìm thấy dòng tiêu đề "Order No" trong file data SCAF.');
    ['reference', 'color', 'size', 'realQty'].forEach(function (k) {
      if (colMap[k] === undefined) throw new Error('Thiếu cột "' + COLS[k][0] + '" trong file data SCAF.');
    });

    var rows = [], last = {};
    for (var rr = headerRow + 1; rr < maxR; rr++) {
      var line = grid[rr] || [];
      var rowTxt = line.map(cellText).join(' ').toLowerCase();
      if (rowTxt.indexOf('total') >= 0 && cellText(line[colMap.size]).trim() === '') break;
      var get = function (k) { return colMap[k] === undefined ? null : line[colMap[k]]; };
      var size = cellText(get('size')).trim();
      if (!size) continue;
      var rec = {};
      ['orderNo', 'story', 'reference', 'kind', 'color'].forEach(function (k) {
        var v = cellText(get(k)).trim();
        rec[k] = v || last[k] || '';
      });
      last = rec;
      rows.push({
        orderNo: rec.orderNo, story: rec.story, reference: rec.reference, kind: rec.kind, color: rec.color,
        sizeRaw: size,
        ctnFrom: num(get('ctnFrom')), ctnTo: num(get('ctnTo')),
        qty: num(get('qty')), boxes: num(get('boxes')), realQty: num(get('realQty')) || num(get('qty')) * num(get('boxes'))
      });
    }
    if (!rows.length) throw new Error('Không đọc được dòng dữ liệu nào trong file data SCAF.');
    return { info: info, rows: rows };
  }

  function groupKey(r) { return r.reference + '||' + r.color; }

  /** Danh sách tổ hợp Reference + Color cần khai báo customs code */
  function listCombos(rows) {
    var seen = {}, out = [];
    rows.forEach(function (r) {
      var k = groupKey(r);
      if (!seen[k]) { seen[k] = 1; out.push({ key: k, orderNo: r.orderNo, reference: r.reference, color: r.color, kind: r.kind }); }
    });
    return out;
  }

  /**
   * mapping[key] = { code, styleNo, hs, price }
   * returns groups: [{code, styleNo, hs, price, poName, lines:[...], summary:[{size, boxes, qty, price}]}]
   */
  function buildGroups(rows, mapping, sizeMode, priceOverrides) {
    priceOverrides = priceOverrides || {};
    var groups = [], byCode = {};
    rows.forEach(function (r) {
      var m = mapping[groupKey(r)] || {};
      var code = (m.code || '').trim();
      if (!code) throw new Error('Chưa nhập Customs code cho ' + r.reference + ' / ' + r.color);
      var g = byCode[code];
      if (!g) {
        g = byCode[code] = { code: code, styleNo: m.styleNo || '', hs: m.hs || '', price: num(m.price), poName: r.orderNo, lines: [], summary: [], _s: {} };
        groups.push(g);
      }
      var size = convertSize(r.sizeRaw, sizeMode);
      g.lines.push({ reference: r.reference, color: r.color, size: size, ctnFrom: r.ctnFrom, ctnTo: r.ctnTo, boxes: r.boxes, qty: r.realQty, perBox: r.qty });
      var sk = String(size);
      if (!g._s[sk]) { g._s[sk] = { size: size, boxes: 0, qty: 0 }; g.summary.push(g._s[sk]); }
      g._s[sk].boxes += r.boxes; g._s[sk].qty += r.realQty;
    });
    groups.forEach(function (g) {
      g.summary.forEach(function (s) {
        var ok = g.code + '-' + s.size;
        s.item = ok;
        s.price = priceOverrides[ok] !== undefined && priceOverrides[ok] !== '' ? num(priceOverrides[ok]) : g.price;
      });
      delete g._s;
    });
    return groups;
  }

  // -------------------------------------------------------------- writer
  var THIN = { style: 'thin' }, MED = { style: 'medium' };
  var RED = 'FFFF0000', BLUE = 'FF0000FF', PURPLE = 'FF7030A0';

  function q(name) { return "'" + name.replace(/'/g, "''") + "'"; }
  function safeSheetName(s) { return String(s).replace(/[\[\]\*\?\/\\:]/g, '_').slice(0, 31); }

  function setC(ws, addr, value, style) {
    var c = ws.getCell(addr);
    c.value = value;
    if (style) {
      if (style.font) c.font = style.font;
      if (style.numFmt) c.numFmt = style.numFmt;
      if (style.alignment) c.alignment = style.alignment;
      if (style.fill) c.fill = style.fill;
      if (style.border) c.border = style.border;
    }
    return c;
  }
  function F(formula, result) { return { formula: formula, result: result }; }
  function boxBorder(ws, r1, c1, r2, c2, edge) {
    edge = edge || THIN;
    for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) {
      var cell = ws.getCell(r, c), b = Object.assign({}, cell.border || {});
      if (r === r1) b.top = edge; if (r === r2) b.bottom = edge;
      if (c === c1) b.left = edge; if (c === c2) b.right = edge;
      cell.border = b;
    }
  }
  function gridBorder(ws, r1, c1, r2, c2) {
    for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++)
      ws.getCell(r, c).border = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  }
  function widths(ws, map) { Object.keys(map).forEach(function (k) { ws.getColumn(k).width = map[k]; }); }
  function fontAll(ws, font) {
    ws.eachRow(function (row) { row.eachCell(function (c) { if (!c.font || !c.font.name) c.font = Object.assign({}, font, c.font || {}); }); });
  }

  /**
   * model = { info:{...header}, groups:[...], settings:{...} }
   */
  function buildWorkbook(ExcelJS, model) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'SCAF Docs Generator';
    wb.calcProperties = { fullCalcOnLoad: true };
    var H = model.header, groups = model.groups;

    // compute totals
    var totalQty = 0, totalBoxes = 0, totalAmt = 0;
    groups.forEach(function (g) {
      g.lines.forEach(function (l) { totalBoxes += l.boxes; });
      g.summary.forEach(function (s) { totalQty += s.qty; totalAmt += s.qty * s.price; });
    });

    // create sheets in the order of the form
    var wsI3 = wb.addWorksheet('INV-3RD', { pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    var wsID = wb.addWorksheet('INV - delta', { pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    var wsPL = wb.addWorksheet('PLIST', { pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    var wsCI = wb.addWorksheet('CI', { pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    groups.forEach(function (g) { g.sheet = safeSheetName(g.code); g.ws = wb.addWorksheet(g.sheet); });

    // ===================================================== Custom code sheets
    groups.forEach(function (g) {
      var ws = g.ws, n = g.lines.length, last = 2 + n;
      widths(ws, { A: 42.3, B: 4, C: 19.4, D: 3, E: 10.7, F: 9, G: 7, H: 3, I: 7, J: 8, K: 11.7, L: 9.6, M: 18, N: 8, O: 11 });
      var red = { name: 'Arial', size: 10, color: { argb: RED } };
      var pur = { name: 'Arial', size: 10, color: { argb: PURPLE } };
      setC(ws, 'A1', g.code, { font: { name: 'Arial', size: 10, bold: true, color: { argb: RED } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } });
      var sumBoxes = 0, sumQty = 0; g.lines.forEach(function (l) { sumBoxes += l.boxes; sumQty += l.qty; });
      setC(ws, 'J1', F('SUM(J3:J' + last + ')', sumBoxes), { font: Object.assign({ bold: true }, red) });
      setC(ws, 'K1', F('SUM(K3:K' + last + ')', sumQty), { font: Object.assign({ bold: true }, red) });
      [['A2', 'PO'], ['C2', 'ref'], ['E2', 'màu'], ['F2', 'size'], ['G2', 'from'], ['I2', 'to'], ['J2', 'kiện'], ['K2', 'số lượng']]
        .forEach(function (x) { setC(ws, x[0], x[1], { font: red }); });
      g.lines.forEach(function (l, i) {
        var r = 3 + i;
        if (i === 0) setC(ws, 'A' + r, g.poName, { font: { name: 'Times New Roman', size: 12, color: { argb: RED } } });
        setC(ws, 'C' + r, l.reference, { font: red });
        setC(ws, 'E' + r, l.color, { font: red });
        setC(ws, 'F' + r, l.size, { font: red, alignment: { horizontal: 'right' } });
        setC(ws, 'G' + r, l.ctnFrom, { font: red });
        setC(ws, 'I' + r, l.ctnTo, { font: red });
        setC(ws, 'J' + r, l.boxes || null, { font: red });
        setC(ws, 'K' + r, l.qty, { font: red });
      });
      // summary by customs code + size
      setC(ws, 'L2', 'size', { font: pur }); setC(ws, 'M2', 'item number', { font: pur });
      setC(ws, 'N2', 'kiện', { font: pur }); setC(ws, 'O2', 'số lượng', { font: pur });
      var sl = 2 + g.summary.length;
      var sb = 0, sq = 0; g.summary.forEach(function (s) { sb += s.boxes; sq += s.qty; });
      setC(ws, 'N1', F('SUM(N3:N' + sl + ')', sb), { font: Object.assign({ bold: true }, pur) });
      setC(ws, 'O1', F('SUM(O3:O' + sl + ')', sq), { font: Object.assign({ bold: true }, pur) });
      g.summary.forEach(function (s, i) {
        var r = 3 + i;
        s.ccRow = r;
        setC(ws, 'L' + r, s.size, { font: pur, alignment: { horizontal: 'right' } });
        setC(ws, 'M' + r, F('$A$1&"-"&L' + r, s.item), { font: pur });
        setC(ws, 'N' + r, F('SUMIF($F$3:$F$' + last + ',L' + r + ',$J$3:$J$' + last + ')', s.boxes), { font: pur });
        setC(ws, 'O' + r, F('SUMIF($F$3:$F$' + last + ',L' + r + ',$K$3:$K$' + last + ')', s.qty), { font: pur });
      });
      ws.views = [{ state: 'frozen', ySplit: 2 }];
    });

    // ===================================================== PLIST
    (function () {
      var ws = wsPL;
      var TNR = { name: 'Times New Roman', size: 10 };
      var B = { name: 'Times New Roman', size: 10, bold: true };
      var redL = { name: 'Times New Roman', size: 10, color: { argb: 'FFC00000' } };
      widths(ws, { A: 5.5, B: 5.5, C: 12, D: 22.4, E: 8, F: 8.5, G: 8.7, H: 11.6, I: 11.3, J: 10, K: 10, L: 21.7 });
      ws.mergeCells('A1:L1'); ws.mergeCells('A2:L2');
      setC(ws, 'A1', 'SCAVI HUE COMPANY', { font: { name: 'Times New Roman', size: 14, bold: true }, alignment: { horizontal: 'center' } });
      setC(ws, 'A2', 'PACKING LIST', { font: { name: 'Times New Roman', size: 14, bold: true }, alignment: { horizontal: 'center' } });
      setC(ws, 'A3', 'DATE:  ', { font: B }); setC(ws, 'C3', F('CI!G4', H.invoiceDate), { font: B, numFmt: 'dd-mmm-yy', alignment: { horizontal: 'left' } });
      setC(ws, 'A4', 'Customers', { font: B }); setC(ws, 'D4', F('CI!A8&""', H.consignee1), { font: B });
      setC(ws, 'A5', 'Supplier', { font: B }); setC(ws, 'D5', 'SCAVI HUE COMPANY', { font: B });
      setC(ws, 'H5', 'Invoice #: ', { font: B }); setC(ws, 'I5', F('CI!G3&""', H.invoiceNo), { font: B });
      setC(ws, 'A6', 'Shipper per', { font: B }); setC(ws, 'D6', H.shipperPer, { font: B }); setC(ws, 'H6', 'Flight', { font: B });
      setC(ws, 'A7', 'Salling on', { font: B }); setC(ws, 'C7', F('CI!B19', H.sailingDate), { font: B, numFmt: 'dd-mmm-yy', alignment: { horizontal: 'left' } });
      setC(ws, 'H7', 'Term of payment :', { font: B });

      var r = 8, hdrRow = 9, first = true, firstData = null;
      groups.forEach(function (g) {
        setC(ws, 'A' + r, F(q(g.sheet) + '!A3', g.poName), { font: { name: 'Times New Roman', size: 12, color: { argb: RED } } });
        r++;
        if (first) {
          var hd = ['CTN\nFrom', 'CTN\nTo', 'Color', 'Reference', 'Size', 'Qty', 'Number\nof Box', 'Real\nQty', 'CBM', 'NW', 'GW', 'CONTAINER'];
          hd.forEach(function (t, i) { var c = ws.getCell(r, i + 1); c.value = t; c.font = B; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
          ws.getRow(r).height = 27; hdrRow = r; r++; first = false;
        }
        g.lines.forEach(function (l, i) {
          var src = 3 + i, sh = q(g.sheet) + '!';
          if (firstData === null) firstData = r;
          setC(ws, 'A' + r, F(sh + 'G' + src, l.ctnFrom), { font: redL, alignment: { horizontal: 'center' } });
          setC(ws, 'B' + r, F(sh + 'I' + src, l.ctnTo), { font: redL, alignment: { horizontal: 'center' } });
          setC(ws, 'C' + r, F(sh + 'E' + src, l.color), { font: { name: 'Arial', size: 10, color: { argb: 'FFC00000' } }, alignment: { horizontal: 'center' } });
          setC(ws, 'D' + r, F(sh + 'C' + src, l.reference), { font: { name: 'Arial', size: 10, color: { argb: 'FFC00000' } }, alignment: { horizontal: 'center' } });
          setC(ws, 'E' + r, F(sh + 'F' + src, l.size), { font: redL, alignment: { horizontal: 'left' } });
          setC(ws, 'F' + r, F('H' + r, l.qty), { font: redL, numFmt: '#,##0' });
          setC(ws, 'G' + r, F(sh + 'J' + src, l.boxes), { font: redL, numFmt: '#,##0' });
          setC(ws, 'H' + r, F(sh + 'K' + src, l.qty), { font: redL, numFmt: '#,##0' });
          r++;
        });
      });
      var lastData = r - 1;
      // 2 empty rows like the form
      r += 2;
      var tot = r;
      setC(ws, 'I' + firstData, H.cbm === '' ? null : num(H.cbm), { font: { name: 'Times New Roman', size: 12 }, numFmt: '0.00' });
      setC(ws, 'J' + firstData, H.nw === '' ? null : num(H.nw), { font: { name: 'Times New Roman', size: 12 }, numFmt: '0.00' });
      setC(ws, 'K' + firstData, H.gw === '' ? null : num(H.gw), { font: { name: 'Times New Roman', size: 12 }, numFmt: '0.00' });
      setC(ws, 'L' + firstData, H.container || null, { font: { name: 'Arial', size: 12, color: { argb: RED } } });
      if (H.containerNote) setC(ws, 'L' + (firstData + 2 <= lastData ? firstData + 2 : firstData + 1), H.containerNote, { font: { name: 'Arial', size: 12, color: { argb: RED } } });
      ['F', 'G', 'H', 'I', 'J', 'K'].forEach(function (c) {
        var res = c === 'G' ? totalBoxes : (c === 'I' ? num(H.cbm) : c === 'J' ? num(H.nw) : c === 'K' ? num(H.gw) : totalQty);
        setC(ws, c + tot, F('SUM(' + c + firstData + ':' + c + (tot - 1) + ')', res), { font: B, numFmt: ['I', 'J', 'K'].indexOf(c) >= 0 ? '0.00' : '0', alignment: { horizontal: 'center' } });
      });
      setC(ws, 'I' + (tot + 1), 'CBM', { font: B, alignment: { horizontal: 'center' } });
      setC(ws, 'J' + (tot + 1), 'KGS', { font: B, alignment: { horizontal: 'center' } });
      setC(ws, 'K' + (tot + 1), 'KGS', { font: B, alignment: { horizontal: 'center' } });
      gridBorder(ws, hdrRow, 1, tot + 1, 12);
      // CBM/NW/GW/CONTAINER columns: one merged-looking block (no inner horizontal lines)
      for (var rr = firstData; rr < tot; rr++) ['I', 'J', 'K', 'L'].forEach(function (c) {
        var cell = ws.getCell(c + rr); cell.border = { left: THIN, right: THIN, top: rr === firstData ? THIN : undefined };
      });
      model._plistTotalRow = tot;
      fontAll(ws, TNR);
    })();

    // ===================================================== CI
    (function () {
      var ws = wsCI;
      var T = function (o) { return Object.assign({ name: 'Times New Roman', size: 12 }, o || {}); };
      widths(ws, { A: 24.6, B: 15.3, C: 13.1, D: 10.9, E: 13.3, F: 12.3, G: 26.9, H: 4, I: 14 });
      ws.mergeCells('A1:G1');
      setC(ws, 'A1', 'COMMERCIAL INVOICE', { font: T({ size: 20, bold: true, underline: true }), alignment: { horizontal: 'center' } });
      setC(ws, 'A3', '1) Shipper/exporter:', { font: T() });
      setC(ws, 'E3', '8) No of Invoice:', { font: T() });
      setC(ws, 'G3', H.invoiceNo, { font: T({ bold: true, color: { argb: RED } }) });
      setC(ws, 'A4', 'SCAVI HUE COMPANY', { font: T({ bold: true }) });
      setC(ws, 'E4', '    Date:', { font: T() });
      setC(ws, 'G4', H.invoiceDate, { font: T({ color: { argb: RED } }), numFmt: 'dd-mmm-yyyy', alignment: { horizontal: 'left' } });
      setC(ws, 'A5', 'Phong Dien Industrial Zone', { font: T() });
      setC(ws, 'A6', 'Phong Dien Ward, Hue City, Viet Nam', { font: T() });
      setC(ws, 'E6', '9) Buyer', { font: T() });
      setC(ws, 'A7', '2)Consignee', { font: T() });
      setC(ws, 'E7', H.buyer1, { font: T({ bold: true }) });
      setC(ws, 'A8', H.consignee1, { font: T({ bold: true }) });
      setC(ws, 'E8', H.buyer2, { font: T({ bold: true }) });
      setC(ws, 'A9', H.consignee2, { font: T() });
      setC(ws, 'A10', H.consignee3, { font: T() });
      setC(ws, 'E10', H.terms, { font: T({ bold: true, color: { argb: BLUE } }) });
      setC(ws, 'A11', '3) Ship to:', { font: T() });
      setC(ws, 'A12', F('A8&""', H.consignee1), { font: T({ bold: true }) });
      setC(ws, 'A13', F('A9&""', H.consignee2), { font: T() });
      setC(ws, 'A14', F('A10&""', H.consignee3), { font: T() });
      setC(ws, 'E15', 'PO: ', { font: T({ bold: true }) });
      setC(ws, 'F15', /^\d+$/.test(H.po) ? Number(H.po) : H.po, { font: T({ color: { argb: RED } }), alignment: { horizontal: 'center' }, border: { top: THIN, left: THIN, bottom: THIN, right: THIN } });
      setC(ws, 'A16', '4) Port of lading:', { font: T() });
      setC(ws, 'B16', '5) Destination:', { font: T() });
      setC(ws, 'A17', H.portLoading, { font: T({ color: { argb: BLUE } }) });
      setC(ws, 'B17', H.destination, { font: T({ color: { argb: BLUE } }) });
      setC(ws, 'A18', '6) Carrier:', { font: T() });
      setC(ws, 'B18', '7) Sailing on:', { font: T() });
      setC(ws, 'E18', 'Country of Origin: Viet Nam', { font: T({ bold: true }) });
      setC(ws, 'A19', H.carrier, { font: T({ color: { argb: BLUE } }), alignment: { horizontal: 'center' } });
      setC(ws, 'B19', H.sailingDate, { font: T({ color: { argb: RED } }), numFmt: 'dd-mmm-yy', alignment: { horizontal: 'center' } });
      setC(ws, 'A20', '12) No of CARTON:', { font: T() });
      setC(ws, 'B20', '13) Description of goods', { font: T() });
      setC(ws, 'E20', '15)          U/P USD', { font: T() });
      setC(ws, 'G20', '16)       Amount  USD', { font: T() });
      setC(ws, 'A21', F('PLIST!G' + model._plistTotalRow, totalBoxes), { font: T({ bold: true, color: { argb: RED } }), alignment: { horizontal: 'center' } });
      setC(ws, 'B21', H.description, { font: T({ color: { argb: RED } }) });
      setC(ws, 'E21', 'FOB', { font: T({ color: { argb: 'FFC00000' } }), alignment: { horizontal: 'center' } });
      setC(ws, 'G21', 'FOB', { font: T(), alignment: { horizontal: 'center' } });
      setC(ws, 'C22', '14) Quantity:', { font: T() });
      setC(ws, 'D22', 'Unit', { font: T() });
      // header boxes
      boxBorder(ws, 3, 1, 6, 4); boxBorder(ws, 7, 1, 10, 4); boxBorder(ws, 11, 1, 15, 4);
      boxBorder(ws, 3, 5, 5, 7); boxBorder(ws, 6, 5, 9, 7); boxBorder(ws, 10, 5, 15, 7);
      boxBorder(ws, 16, 1, 17, 1); boxBorder(ws, 16, 2, 17, 4); boxBorder(ws, 18, 1, 19, 1); boxBorder(ws, 18, 2, 19, 4);
      boxBorder(ws, 16, 5, 19, 7); boxBorder(ws, 20, 1, 21, 1); boxBorder(ws, 20, 2, 21, 4);
      boxBorder(ws, 20, 5, 21, 6); boxBorder(ws, 20, 7, 21, 7); boxBorder(ws, 22, 3, 22, 4);

      setC(ws, 'A24', 'Season : ' + (H.season || ''), { font: T({ bold: true, underline: true }) });
      var r = 25, lineRows = [];
      groups.forEach(function (g, gi) {
        setC(ws, 'A' + r, 'ORDER NO', { font: T({ bold: true, underline: true }) });
        setC(ws, 'B' + r, g.styleNo, { font: T() });
        g.ciOrderRow = r;
        r++;
        setC(ws, 'B' + r, 'HS CODE', { font: T() });
        r++;
        g.summary.forEach(function (s) {
          var sh = q(g.sheet) + '!';
          var hs = /^\d+$/.test(String(g.hs)) ? Number(g.hs) : g.hs;
          setC(ws, 'A' + r, F(sh + 'M' + s.ccRow, s.item), { font: T({ color: { argb: RED } }) });
          setC(ws, 'B' + r, hs, { font: T(), alignment: { horizontal: 'left' } });
          setC(ws, 'C' + r, F(sh + 'O' + s.ccRow, s.qty), { font: T({ color: { argb: RED } }), numFmt: '#,##0', alignment: { horizontal: 'center' } });
          setC(ws, 'D' + r, 'PCE', { font: T() });
          setC(ws, 'E' + r, s.price, { font: T({ color: { argb: RED } }), numFmt: '0.00000', alignment: { horizontal: 'center' } });
          setC(ws, 'G' + r, F('E' + r + '*C' + r, s.qty * s.price), { font: T(), numFmt: '#,##0.00' });
          s.ciRow = r; lineRows.push(r);
          r++;
        });
        r++;
      });
      var first = 27, lastLine = r - 1;
      r++;
      var tot = r;
      setC(ws, 'C' + tot, F('SUM(C' + first + ':C' + lastLine + ')', totalQty), { font: T({ bold: true }), numFmt: '#,##0', alignment: { horizontal: 'right' } });
      setC(ws, 'G' + tot, F('SUM(G' + first + ':G' + lastLine + ')', totalAmt), { font: T({ bold: true }), numFmt: '#,##0.00', alignment: { horizontal: 'right' } });
      setC(ws, 'C' + (tot + 1), 'PCS', { font: T({ bold: true }), alignment: { horizontal: 'right' } });
      setC(ws, 'G' + (tot + 1), 'USD', { font: T({ bold: true }), alignment: { horizontal: 'right' } });
      boxBorder(ws, tot, 3, tot + 1, 3); boxBorder(ws, tot, 7, tot + 1, 7);

      // HS code summary
      r = tot + 3;
      ['H.S CODE', 'QUANTITY', 'FOB VALUE'].forEach(function (t, i) { ws.getCell(r, 2 + i).value = t; ws.getCell(r, 2 + i).font = T({ bold: true, size: 10 }); });
      var hsList = [];
      groups.forEach(function (g) { var h = /^\d+$/.test(String(g.hs)) ? Number(g.hs) : g.hs; if (hsList.indexOf(h) < 0) hsList.push(h); });
      var hsStart = r + 1;
      hsList.forEach(function (h, i) {
        var rr = hsStart + i, q1 = 0, v1 = 0;
        groups.forEach(function (g) { var gh = /^\d+$/.test(String(g.hs)) ? Number(g.hs) : g.hs; if (gh === h) g.summary.forEach(function (s) { q1 += s.qty; v1 += s.qty * s.price; }); });
        setC(ws, 'B' + rr, h, { font: T({ size: 10 }), alignment: { horizontal: 'left' } });
        setC(ws, 'C' + rr, F('SUMIF($B$' + first + ':$B$' + lastLine + ',B' + rr + ',$C$' + first + ':$C$' + lastLine + ')', q1), { font: T({ size: 10 }), numFmt: '#,##0' });
        setC(ws, 'D' + rr, F('SUMIF($B$' + first + ':$B$' + lastLine + ',B' + rr + ',$G$' + first + ':$G$' + lastLine + ')', v1), { font: T({ size: 10 }), numFmt: '#,##0.00' });
      });
      var hsTot = hsStart + hsList.length;
      setC(ws, 'B' + hsTot, 'TOTAL:', { font: T({ bold: true, size: 10 }) });
      setC(ws, 'C' + hsTot, F('SUM(C' + hsStart + ':C' + (hsTot - 1) + ')', totalQty), { font: T({ bold: true, size: 10, color: { argb: RED } }), numFmt: '#,##0' });
      setC(ws, 'D' + hsTot, F('SUM(D' + hsStart + ':D' + (hsTot - 1) + ')', totalAmt), { font: T({ bold: true, size: 10, color: { argb: RED } }), numFmt: '#,##0.00' });
      gridBorder(ws, r, 2, hsTot, 4);

      r = hsTot + 2;
      ws.mergeCells('A' + r + ':G' + r);
      setC(ws, 'A' + r, 'ON BEHALF OF SCAVI HUE COMPANY', { font: T({ bold: true, underline: true }), alignment: { horizontal: 'center' } });
      fontAll(ws, T());
    })();

    // ===================================================== INV-3RD & INV - delta
    function invHeader(ws, isDelta) {
      var A = function (o) { return Object.assign({ name: 'Arial', size: 10 }, o || {}); };
      var lastCol = isDelta ? 'O' : 'P';
      ws.mergeCells('A1:' + lastCol + '1');
      setC(ws, 'A1', 'COMMERCIAL INVOICE', { font: A({ size: 24, bold: true, underline: true }), alignment: { horizontal: 'center' } });
      setC(ws, (isDelta ? 'O' : 'P') + '2', '1 of 1', { font: A({ bold: true, size: 8 }) });
      setC(ws, 'A3', isDelta ? 'Supplier :' : 'Seller', { font: A({ bold: true }) });
      setC(ws, 'G3', 'Invoice No. and Date :', { font: A({ bold: true }) });
      if (isDelta) {
        setC(ws, 'B4', F('CI!A4&""', 'SCAVI HUE COMPANY'), { font: A({ bold: true }) });
        setC(ws, 'B5', F('CI!A5&""', 'Phong Dien Industrial Zone'), { font: A() });
        setC(ws, 'B6', F('CI!A6&""', 'Phong Dien Ward, Hue City, Viet Nam'), { font: A() });
        setC(ws, 'G6', 'Buyer: ', { font: A({ bold: true }) });
        setC(ws, 'G7', F('CI!E7&""', H.buyer1), { font: A() });
        setC(ws, 'G8', F('CI!E8&""', H.buyer2), { font: A() });
        if (H.buyerTel) setC(ws, 'G9', H.buyerTel, { font: A({ bold: true }) });
        setC(ws, 'G12', 'L/C Issuing Bank :', { font: A({ bold: true }) });
        setC(ws, 'N4', F('CI!G4', H.invoiceDate), { font: A(), numFmt: 'dd-mmm-yyyy' });
        ws.mergeCells('N4:O4');
      } else {
        setC(ws, 'B4', F('CI!E7&""', H.buyer1), { font: A({ bold: true }) });
        setC(ws, 'B5', F('CI!E8&""', H.buyer2), { font: A() });
        setC(ws, 'G6', 'BUYER', { font: A({ bold: true }) });
        setC(ws, 'G7', F('B11&""', H.consignee1), { font: A({ bold: true }) });
        setC(ws, 'G8', F('B12&""', H.consignee2), { font: A() });
        setC(ws, 'G9', F('B13&""', H.consignee3), { font: A() });
        setC(ws, 'M4', F('CI!G4', H.invoiceDate), { font: A(), numFmt: 'dd-mmm-yyyy' });
        ws.mergeCells('M4:P4');
      }
      setC(ws, 'H4', F('CI!G3&""', H.invoiceNo), { font: A() });
      setC(ws, 'L4', 'DATE:', { font: A() });
      setC(ws, 'A10', 'CONSIGNEE:', { font: A({ bold: true }) });
      setC(ws, 'B11', F('CI!A8&""', H.consignee1), { font: A({ bold: true }) });
      setC(ws, 'B12', F('CI!A9&""', H.consignee2), { font: A() });
      setC(ws, 'B13', F('CI!A10&""', H.consignee3), { font: A() });
      if (isDelta && H.consignee4) setC(ws, 'B14', H.consignee4, { font: A() });
      setC(ws, 'G14', '  Terms of Sale : ', { font: A({ bold: true }) });
      setC(ws, 'G15', F('CI!E10&""', H.terms), { font: A() });
      setC(ws, 'A17', 'SHIP TO :', { font: A({ bold: true }) });
      if (!isDelta) { setC(ws, 'H17', 'CONTAINER', { font: A({ bold: true }) }); setC(ws, 'J17', 'SEAL', { font: A({ bold: true }) }); }
      setC(ws, 'B18', F('B11&""', H.consignee1), { font: A({ bold: true }) });
      setC(ws, 'B19', F('B12&""', H.consignee2), { font: A() });
      setC(ws, 'B20', F('B13&""', H.consignee3), { font: A() });
      if (isDelta && H.consignee4) setC(ws, 'B21', F('B14&""', H.consignee4), { font: A() });
      if (isDelta) {
        setC(ws, 'H23', 'COUNTRY OF ORIGIN : VIETNAM', { font: A() });
        setC(ws, 'H24', F('"MODE OF TRANSPORTATION : BY "&CI!A19', 'MODE OF TRANSPORTATION : BY ' + H.carrier), { font: A() });
        setC(ws, 'H25', 'FREIGHT                      : COLLECT', { font: A() });
      } else {
        setC(ws, 'H22', 'B/L NO:', { font: A() });
        setC(ws, 'H23', 'COUNTRY OF ORIGIN :', { font: A() }); setC(ws, 'K23', 'VIETNAM', { font: A() });
        setC(ws, 'H24', 'MODE OF TRANSPORTATION :', { font: A() }); setC(ws, 'K24', 'BY ' + (H.shipperPer || 'SEA'), { font: A() });
        setC(ws, 'H25', 'FREIGHT                      :', { font: A() }); setC(ws, 'K25', 'COLLECT', { font: A() });
      }
      setC(ws, 'A26', 'Port of Loading :', { font: A({ bold: true }) });
      setC(ws, 'D26', 'Final Destination :', { font: A({ bold: true }) });
      setC(ws, 'H26', 'TERM OF DELIVERY    :', { font: A() });
      setC(ws, 'K26', H.termDelivery, { font: A() });
      ws.mergeCells('A27:C27'); ws.mergeCells('D27:F27');
      setC(ws, 'A27', F('CI!A17&""', H.portLoading), { font: A() });
      setC(ws, 'D27', F('CI!B17&""', H.destination), { font: A(), alignment: { horizontal: 'center' } });
      setC(ws, 'H27', 'TOTAL CARTON           :', { font: A() });
      setC(ws, 'H28', 'TOTAL NET WEIGHT   :', { font: A() });
      setC(ws, 'H29', 'TOTAL GROSS  WEIGHT :', { font: A() });
      setC(ws, 'H30', 'TOTAL MEASUREMENT  :', { font: A() });
      ['27', '28', '29', '30'].forEach(function (rr) { ws.mergeCells('K' + rr + ':L' + rr); });
      var pt = model._plistTotalRow;
      if (isDelta) {
        setC(ws, 'K27', F('PLIST!G' + pt, totalBoxes), { font: A(), numFmt: '0 "CTNS"' });
        setC(ws, 'K28', F('PLIST!J' + pt, num(H.nw)), { font: A(), numFmt: '#,##0.00 "KGS"' });
        setC(ws, 'K29', F('PLIST!K' + pt, num(H.gw)), { font: A(), numFmt: '#,##0.00 "KGS"' });
        setC(ws, 'K30', F('PLIST!I' + pt, num(H.cbm)), { font: A(), numFmt: '0.00 "CBM"' });
      } else {
        setC(ws, 'K27', F("'INV - delta'!K27", totalBoxes), { font: A(), numFmt: '0 "CTNS"' });
        setC(ws, 'K28', F("'INV - delta'!K28", num(H.nw)), { font: A(), numFmt: '#,##0.00 "KGS"' });
        setC(ws, 'K29', F("'INV - delta'!K29", num(H.gw)), { font: A(), numFmt: '#,##0.00 "KGS"' });
        setC(ws, 'K30', F("'INV - delta'!K30", num(H.cbm)), { font: A(), numFmt: '0.00 "CBM"' });
      }
      setC(ws, 'A28', 'Vessel/Flight :', { font: A({ bold: true }) });
      ws.mergeCells('A29:C29'); ws.mergeCells('D29:F29');
      setC(ws, 'A29', F('CI!A19&""', H.carrier), { font: A(), alignment: { horizontal: 'center' } });
      setC(ws, 'D29', F('CI!B19', H.sailingDate), { font: A(), numFmt: 'dd-mmm-yyyy', alignment: { horizontal: 'center' } });
      // frame lines
      var lc = isDelta ? 15 : 16;
      boxBorder(ws, 1, 1, 2, lc);
      boxBorder(ws, 3, 1, 9, 6); boxBorder(ws, 10, 1, 16, 6); boxBorder(ws, 17, 1, 25, 6);
      boxBorder(ws, 3, 7, 5, lc); boxBorder(ws, 6, 7, 11, lc); boxBorder(ws, 12, 7, 13, lc); boxBorder(ws, 14, 7, 16, lc);
      boxBorder(ws, 17, 7, 31, lc);
      boxBorder(ws, 26, 1, 27, 3); boxBorder(ws, 26, 4, 27, 6); boxBorder(ws, 28, 1, 31, 3); boxBorder(ws, 28, 4, 31, 6);
    }

    function invBody(ws, isDelta) {
      var A = function (o) { return Object.assign({ name: 'Arial', size: 10 }, o || {}); };
      var redA = A({ color: { argb: RED } });
      var lc = isDelta ? 15 : 16;
      setC(ws, 'D32', 'Description of Goods', { font: A({ bold: true }) });
      setC(ws, 'J32', 'Quantity/unit', { font: A({ bold: true }) });
      if (isDelta) {
        ws.mergeCells('L32:N32');
        setC(ws, 'L32', 'Unit Price', { font: A({ bold: true }), alignment: { horizontal: 'center' } });
        setC(ws, 'O32', 'Amount', { font: A({ bold: true }), alignment: { horizontal: 'center' } });
        setC(ws, 'L33', 'FOB', { font: A({ bold: true, color: { argb: RED } }), alignment: { horizontal: 'center' } });
        setC(ws, 'N33', 'FOB', { font: A({ bold: true }), alignment: { horizontal: 'center' } });
        setC(ws, 'O33', 'FOB', { font: A({ bold: true }), alignment: { horizontal: 'center' } });
        setC(ws, 'O35', F("'INV-3RD'!P35&\"\"", H.termDelivery), { font: A({ bold: true, underline: true }), alignment: { horizontal: 'right' } });
        [['D36', 'ORDER NO'], ['E36', 'STYLE CODE'], ['F36', 'ITEM NUMBER'], ['H36', 'HS CODE '], ['K36', "Q'ty / PC"], ['L36', 'Price/PC'], ['M36', "Q'ty / set"], ['N36', 'Price/set']]
          .forEach(function (x) { setC(ws, x[0], x[1], { font: A({ bold: true }), alignment: { horizontal: 'center' } }); });
      } else {
        ws.mergeCells('L32:M32'); ws.mergeCells('N32:O32');
        setC(ws, 'L32', 'Unit Price', { font: A({ bold: true }), alignment: { horizontal: 'center' } });
        setC(ws, 'N32', 'Unit Price', { font: A({ bold: true }), alignment: { horizontal: 'center' } });
        setC(ws, 'P32', 'Amount', { font: A({ bold: true }), alignment: { horizontal: 'center' } });
        ['L33', 'M33', 'N33', 'O33', 'P33'].forEach(function (a) { setC(ws, a, 'FOB', { font: A({ bold: true }), alignment: { horizontal: 'center' } }); });
        [['D35', 'ORDER NO'], ['E35', 'STYLE CODE'], ['F35', 'ITEM NUMBER'], ['G35', 'INDIVIDUAL\nSTYLE NO'], ['H35', 'SET STYLE NO'], ['J35', 'HS CODE ']]
          .forEach(function (x) { setC(ws, x[0], x[1], { font: A({ bold: true }), alignment: { horizontal: 'center', vertical: 'middle', wrapText: true } }); });
        setC(ws, 'P35', H.termDelivery, { font: A({ bold: true, underline: true }), alignment: { horizontal: 'right' } });
      }
      setC(ws, 'D34', F('CI!B21&""', H.description), { font: A({ bold: true, underline: true }) });
      boxBorder(ws, 32, 1, 33, 3); boxBorder(ws, 32, 4, 33, 9); boxBorder(ws, 32, 10, 33, 11); boxBorder(ws, 32, 12, 33, lc);

      var r = isDelta ? 38 : 38;
      var brand = 'BRAND: ' + (H.brand || '');
      var itemFirst = null, itemLast = null;
      if (isDelta) {
        setC(ws, 'B38', F('B41', brand), { font: A({ bold: true }) });
        r = 39;
      }
      groups.forEach(function (g, gi) {
        if (isDelta) {
          setC(ws, 'B' + r, F(q(g.sheet) + '!A3', g.poName), { font: A({ bold: true }) });
          r += 2;
          var hr = r;
          setC(ws, 'B' + hr, gi === 0 ? F("'INV-3RD'!B" + (g.i3BrandRow || 40), brand) : brand, { font: A({ bold: true }) });
          setC(ws, 'D' + hr, F('CI!B' + g.ciOrderRow, g.styleNo), { font: redA });
          setC(ws, 'E' + hr, F('LEFT(F' + (hr + 1) + ',6)', String(g.summary[0].item).slice(0, 6)), { font: redA });
          setC(ws, 'P' + hr, num(H.pcsPerSet) || 1, { font: A({ color: { argb: 'FFFFFFFF' } }) });
          r++;
          g.summary.forEach(function (s) {
            if (itemFirst === null) itemFirst = r;
            var per = num(H.pcsPerSet) || 1;
            setC(ws, 'F' + r, F('CI!A' + s.ciRow, s.item), { font: redA, alignment: { horizontal: 'center' } });
            setC(ws, 'H' + r, F('CI!B' + s.ciRow, /^\d+$/.test(String(g.hs)) ? Number(g.hs) : g.hs), { font: redA, alignment: { horizontal: 'center' } });
            setC(ws, 'K' + r, F('CI!C' + s.ciRow, s.qty), { font: redA, numFmt: '#,##0 "PCS"' });
            setC(ws, 'L' + r, F('CI!E' + s.ciRow, s.price), { font: redA, numFmt: '"$"#,##0.0000' });
            setC(ws, 'M' + r, F('K' + r + '/$P$' + hr, s.qty / per), { font: redA, numFmt: '#,##0.0 "SETS"' });
            setC(ws, 'N' + r, F('L' + r + '*$P$' + hr, s.price * per), { font: redA, numFmt: '"$"#,##0.000' });
            setC(ws, 'O' + r, F('L' + r + '*K' + r, s.qty * s.price), { font: redA, numFmt: '"$"#,##0.00' });
            itemLast = r; r++;
          });
          r++;
        } else {
          setC(ws, 'B' + r, F("'INV - delta'!B" + (39 + 0), g.poName), { font: A({ bold: true }) });
          if (gi > 0) ws.getCell('B' + r).value = F(q(g.sheet) + '!A3', g.poName);
          r += 2;
          setC(ws, 'B' + r, brand, { font: A({ bold: true }) });
          g.i3BrandRow = r;
          r++;
          var hr2 = r;
          setC(ws, 'D' + hr2, F('CI!B' + g.ciOrderRow, g.styleNo), { font: redA });
          setC(ws, 'E' + hr2, F('LEFT(F' + (hr2 + 1) + ',6)', String(g.summary[0].item).slice(0, 6)), { font: redA });
          r++;
          g.summary.forEach(function (s) {
            if (itemFirst === null) itemFirst = r;
            setC(ws, 'F' + r, F('CI!A' + s.ciRow, s.item), { font: redA, alignment: { horizontal: 'center' } });
            setC(ws, 'J' + r, F('CI!B' + s.ciRow, /^\d+$/.test(String(g.hs)) ? Number(g.hs) : g.hs), { font: redA, alignment: { horizontal: 'center' } });
            setC(ws, 'L' + r, F('CI!C' + s.ciRow, s.qty), { font: redA, numFmt: '#,##0 "PCS"' });
            setC(ws, 'M' + r, F('CI!E' + s.ciRow, s.price), { font: redA, numFmt: '"$"#,##0.0000' });
            setC(ws, 'P' + r, F('M' + r + '*L' + r, s.qty * s.price), { font: redA, numFmt: '"$"#,##0.00' });
            itemLast = r; r++;
          });
          r++;
        }
      });
      r++;
      var tot = r, B = A({ bold: true });
      setC(ws, 'D' + tot, 'TOTAL ', { font: B, alignment: { horizontal: 'center' } });
      if (isDelta) {
        setC(ws, 'K' + tot, F('SUM(K' + itemFirst + ':K' + itemLast + ')', totalQty), { font: B, numFmt: '#,##0 "PCS"' });
        setC(ws, 'M' + tot, F('SUM(M' + itemFirst + ':M' + itemLast + ')', totalQty / (num(H.pcsPerSet) || 1)), { font: B, numFmt: '#,##0.0 "SETS"' });
        setC(ws, 'O' + tot, F('SUM(O' + itemFirst + ':O' + itemLast + ')', totalAmt), { font: B, numFmt: '"US$"#,##0.00' });
        for (var c = 4; c <= lc; c++) { ws.getCell(tot, c).border = { top: THIN, bottom: MED }; }
        var gt = tot + 3;
        ws.mergeCells('D' + gt + ':E' + gt);
        setC(ws, 'D' + gt, 'G.TOTAL', { font: { name: 'Cambria', size: 11, bold: true }, alignment: { horizontal: 'center' } });
        setC(ws, 'K' + gt, F('K' + tot, totalQty), { font: B, numFmt: '#,##0 "PCS"' });
        setC(ws, 'M' + gt, F('M' + tot, totalQty / (num(H.pcsPerSet) || 1)), { font: B, numFmt: '#,##0.0 "SETS"' });
        setC(ws, 'O' + gt, F('O' + tot, totalAmt), { font: B, numFmt: '"$"#,##0.00' });
        for (var c2 = 4; c2 <= lc; c2++) { ws.getCell(gt, c2).border = { top: MED, bottom: MED }; }
        setC(ws, 'J' + (gt + 2), 'Signed by', { font: B });
        setC(ws, 'D' + (gt + 7), "By virtue of the clause of property reserve, the goods relative to the present orders will remain SCAVI's property up to their complete payment ", { font: A({ bold: true, size: 8 }) });
        boxBorder(ws, 34, 1, gt + 9, lc);
      } else {
        setC(ws, 'L' + tot, F('SUM(L' + itemFirst + ':L' + itemLast + ')', totalQty), { font: B, numFmt: '#,##0 "PCS"' });
        setC(ws, 'P' + tot, F('SUM(P' + itemFirst + ':P' + itemLast + ')', totalAmt), { font: B, numFmt: '"US$"#,##0.00' });
        for (var c3 = 4; c3 <= lc; c3++) { ws.getCell(tot, c3).border = { top: THIN, bottom: MED }; }
        setC(ws, 'J' + (tot + 3), 'Signed by', { font: B });
        boxBorder(ws, 34, 1, tot + 9, lc);
      }
    }

    // INV-3RD first (INV - delta refers to its brand row), then delta
    widths(wsI3, { A: 1, B: 20.4, C: 12, D: 14.3, E: 14.7, F: 17, G: 15.7, H: 13.6, I: 3, J: 14.3, K: 5.7, L: 12.3, M: 11.7, N: 6.3, O: 6.6, P: 16 });
    widths(wsID, { A: 1, B: 14, C: 12, D: 12.6, E: 12, F: 17, G: 2.7, H: 13, I: 3, J: 6, K: 12.4, L: 11.6, M: 13.6, N: 11.3, O: 14.3, P: 4 });
    invHeader(wsI3, false); invBody(wsI3, false);
    invHeader(wsID, true); invBody(wsID, true);
    wsI3.getCell('K26').value = F('P35', H.termDelivery);
    wsID.getCell('K26').value = F('O35', H.termDelivery);
    fontAll(wsI3, { name: 'Arial', size: 10 }); fontAll(wsID, { name: 'Arial', size: 10 });

    wb.views = [{ activeTab: 3 }];
    return wb;
  }

  function suggestFileName(H, groups) {
    var po = groups[0] ? groups[0].poName : '';
    return (H.invoiceNo + ' INV ' + po).replace(/[\\\/:*?"<>|]+/g, ' ').trim() + '.xlsx';
  }

  return {
    parseScafSheet: parseScafSheet, listCombos: listCombos, buildGroups: buildGroups,
    buildWorkbook: buildWorkbook, convertSize: convertSize, parseDate: parseDate, fmtDate: fmtDate,
    suggestFileName: suggestFileName, cellText: cellText
  };
});
