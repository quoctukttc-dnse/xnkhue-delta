(function () {
  'use strict';
  var C = window.ScafCore;
  var LS_MAP = 'scafdocs.mapping.v1', LS_HDR = 'scafdocs.header.v1', LS_SIZE = 'scafdocs.sizeMode.v1';

  function load(k, d) { try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
  function save(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { } }
  var $ = function (s) { return document.querySelector(s); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  // header fields: key, label, default, type, remember(true = nhớ giá trị cho lần sau)
  var FIELDS = [
    ['invoiceNo', 'Số Invoice (SCAF)', '', 'text', false],
    ['invoiceDate', 'Ngày Invoice', '', 'date', false],
    ['sailingDate', 'Sailing on (ngày tàu chạy)', '', 'date', false],
    ['po', 'PO', '', 'text', false],
    ['carrier', 'Carrier / Vessel', '', 'text', false],
    ['shipperPer', 'Shipper per', 'SEA', 'text', true],
    ['portLoading', 'Port of loading', 'DA NANG PORT', 'text', true],
    ['destination', 'Destination', 'INCHEON, KOREA', 'text', true],
    ['termDelivery', 'Term of delivery', 'FOB DANANG - VIETNAM', 'text', true],
    ['terms', 'Terms of payment', 'TT, 30 days from B/L', 'text', true],
    ['description', 'Description of goods', 'BRIEF', 'text', true],
    ['brand', 'Brand', 'ADIDAS', 'text', true],
    ['season', 'Season', 'S24', 'text', true],
    ['pcsPerSet', 'Số PCS / set', '1', 'number', true],
    ['buyer1', 'Buyer – dòng 1', '', 'text', true],
    ['buyer2', 'Buyer – dòng 2', '', 'text', true],
    ['buyerTel', 'Buyer TEL (INV - delta)', 'TEL: 41 79 586 7403', 'text', true],
    ['consignee1', 'Consignee – tên', 'Pond Group Co., Ltd.', 'text', true],
    ['consignee2', 'Consignee – địa chỉ 1', '42F, Tower1 Parc1, 108, Yeoui-daero, Yeongdeungpo-gu', 'text', true],
    ['consignee3', 'Consignee – địa chỉ 2', 'Seoul, Republic of Korea', 'text', true],
    ['consignee4', 'Consignee – dòng 4 (INV - delta)', 'REPUBLIC OF KOREA (07335)', 'text', true],
    ['cbm', 'Tổng CBM', '', 'number', false],
    ['nw', 'Tổng NW (kg)', '', 'number', false],
    ['gw', 'Tổng GW (kg)', '', 'number', false],
    ['container', 'Container No.', '', 'text', false],
    ['containerNote', 'Ghi chú container', 'A PART OF 1*40HC', 'text', true]
  ];

  var state = { data: null, mapping: load(LS_MAP, {}), priceOv: {}, groups: [] };

  // -------------------------------------------------- header form
  function renderHeader(info) {
    var saved = load(LS_HDR, {});
    var html = FIELDS.map(function (f) {
      var v = f[4] && saved[f[0]] !== undefined ? saved[f[0]] : f[2];
      if (info) {
        if (f[0] === 'invoiceNo' && info.invoiceNo) v = info.invoiceNo;
        if (f[0] === 'invoiceDate' && info.invoiceDate) v = C.fmtDate(info.invoiceDate);
        if (f[0] === 'sailingDate' && (info.sailingDate || info.invoiceDate)) v = C.fmtDate(info.sailingDate || info.invoiceDate);
        if (f[0] === 'po' && info.po) v = info.po;
        if (f[0] === 'buyer1' && info.buyer1) v = info.buyer1;
        if (f[0] === 'buyer2' && info.buyer2) v = info.buyer2;
        if (f[0] === 'portLoading' && info.from && !saved.portLoading) v = /port/i.test(info.from) ? info.from : info.from + ' PORT';
        if (f[0] === 'shipperPer' && info.shipperPer) v = info.shipperPer;
      }
      return '<label>' + esc(f[1]) + '<input type="' + f[3] + '" step="any" data-h="' + f[0] + '" value="' + esc(v) + '"></label>';
    }).join('');
    $('#hdr').innerHTML = html;
  }
  function readHeader() {
    var H = {}, rem = load(LS_HDR, {});
    FIELDS.forEach(function (f) {
      var v = $('[data-h="' + f[0] + '"]').value.trim();
      if (f[4]) rem[f[0]] = v;
      if (f[3] === 'date') v = v ? new Date(v + 'T00:00:00Z') : null;
      H[f[0]] = v;
    });
    save(LS_HDR, rem);
    return H;
  }

  // -------------------------------------------------- mapping table
  function renderMapping() {
    var combos = C.listCombos(state.data.rows);
    var tb = $('#mapTbl tbody');
    tb.innerHTML = combos.map(function (c) {
      var m = state.mapping[c.key] || {};
      var code = m.code || '';
      var sug = (c.reference.split('-').pop() || '') + '-';
      return '<tr data-k="' + esc(c.key) + '">' +
        '<td>' + esc(c.orderNo) + '</td><td>' + esc(c.reference) + '</td><td>' + esc(c.color) + '</td>' +
        '<td><input type="text" data-f="code" placeholder="vd: ' + esc(sug) + '977" value="' + esc(code) + '"' + (code ? '' : ' class="miss"') + '></td>' +
        '<td><input type="text" data-f="styleNo" value="' + esc(m.styleNo || '') + '"></td>' +
        '<td><input type="text" data-f="hs" value="' + esc(m.hs || '') + '"></td>' +
        '<td><input type="number" step="any" data-f="price" value="' + esc(m.price == null ? '' : m.price) + '"></td></tr>';
    }).join('');
  }
  function readMapping() {
    $('#mapTbl tbody').querySelectorAll('tr').forEach(function (tr) {
      var k = tr.getAttribute('data-k'), m = state.mapping[k] || {};
      tr.querySelectorAll('input').forEach(function (inp) {
        var f = inp.getAttribute('data-f'), v = inp.value.trim();
        if (f === 'code') { v = v.replace(/-$/, ''); inp.classList.toggle('miss', !/^.+-.+$/.test(v)); }
        m[f] = v;
      });
      state.mapping[k] = m;
    });
    save(LS_MAP, state.mapping);
  }

  function sizeMode() { return document.querySelector('input[name=sizeMode]:checked').value; }

  // -------------------------------------------------- summary
  function fmt(n, d) { return Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d }); }
  function recompute() {
    if (!state.data) return;
    readMapping();
    save(LS_SIZE, sizeMode());
    var err = $('#err');
    err.hidden = true;
    try {
      state.groups = C.buildGroups(state.data.rows, state.mapping, sizeMode(), state.priceOv);
    } catch (e) {
      state.groups = [];
      $('#sumTbl tbody').innerHTML = '<tr><td colspan="6" style="color:var(--mute)">' + esc(e.message) + '</td></tr>';
      $('#sumTbl tfoot').innerHTML = '';
      $('#go').disabled = true;
      return;
    }
    var tb = '', tq = 0, tbx = 0, ta = 0;
    state.groups.forEach(function (g) {
      g.summary.forEach(function (s) {
        tq += s.qty; tbx += s.boxes; ta += s.qty * s.price;
        tb += '<tr><td>' + esc(s.item) + '</td><td>' + esc(g.hs) + '</td><td class="r">' + fmt(s.boxes, 0) + '</td><td class="r">' + fmt(s.qty, 0) +
          '</td><td class="r"><input type="number" step="any" style="max-width:110px;text-align:right" data-item="' + esc(s.item) + '" value="' + (s.price || '') + '"></td><td class="r">' + fmt(s.qty * s.price, 2) + '</td></tr>';
      });
    });
    $('#sumTbl tbody').innerHTML = tb;
    $('#sumTbl tfoot').innerHTML = '<tr><td>TỔNG</td><td></td><td class="r">' + fmt(tbx, 0) + '</td><td class="r">' + fmt(tq, 0) + '</td><td></td><td class="r">' + fmt(ta, 2) + '</td></tr>';
    $('#sumTbl tbody').querySelectorAll('input[data-item]').forEach(function (inp) {
      inp.addEventListener('change', function () { state.priceOv[inp.getAttribute('data-item')] = inp.value; recompute(); });
    });
    var H0 = { invoiceNo: $('[data-h="invoiceNo"]').value };
    $('#fname').value = C.suggestFileName(H0, state.groups);
    $('#go').disabled = false;
  }

  // -------------------------------------------------- file load
  function showMsg(t, isErr) { var m = $('#msg'); m.hidden = false; m.className = 'msg' + (isErr ? ' err' : ''); m.textContent = t; }
  async function handleFile(file) {
    if (!file) return;
    try {
      var buf = await file.arrayBuffer();
      var wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      var parsed = null, lastErr = null;
      for (var i = 0; i < wb.worksheets.length && !parsed; i++) {
        try { parsed = C.parseScafSheet(wb.worksheets[i]); } catch (e) { lastErr = e; }
      }
      if (!parsed) throw lastErr || new Error('Không đọc được file.');
      var po = (parsed.rows[0].orderNo.match(/^\s*(\d+)/) || [])[1] || '';
      parsed.info.po = po;
      state.data = parsed; state.priceOv = {};
      $('#dropText').innerHTML = '<b>' + esc(file.name) + '</b> — bấm để chọn file khác';
      var tq = parsed.rows.reduce(function (a, r) { return a + r.realQty; }, 0);
      showMsg('Đã đọc ' + parsed.rows.length + ' dòng, tổng ' + tq.toLocaleString('en-US') + ' pcs. Invoice: ' + (parsed.info.invoiceNo || '?'));
      renderHeader(parsed.info);
      renderMapping();
      $('#after').hidden = false;
      recompute();
    } catch (e) {
      console.error(e);
      showMsg('Lỗi: ' + e.message, true);
    }
  }

  // -------------------------------------------------- export
  async function exportXlsx() {
    try {
      recompute();
      if (!state.groups.length) return;
      var H = readHeader();
      var bad = state.groups.filter(function (g) { return g.summary.some(function (s) { return !s.price; }); });
      if (bad.length && !confirm('Còn dòng chưa có đơn giá (' + bad.map(function (g) { return g.code; }).join(', ') + '). Vẫn xuất file?')) return;
      var wb = C.buildWorkbook(ExcelJS, { header: H, groups: state.groups });
      var out = await wb.xlsx.writeBuffer();
      var blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var a = document.createElement('a');
      var name = $('#fname').value.trim() || 'chung-tu.xlsx';
      if (!/\.xlsx$/i.test(name)) name += '.xlsx';
      a.href = URL.createObjectURL(blob); a.download = name;
      document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    } catch (e) {
      console.error(e);
      var err = $('#err'); err.hidden = false; err.textContent = 'Lỗi khi xuất file: ' + e.message;
    }
  }

  // -------------------------------------------------- wire up
  var drop = $('#drop');
  $('#file').addEventListener('change', function (e) { handleFile(e.target.files[0]); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { handleFile(e.dataTransfer.files[0]); });
  $('#mapTbl').addEventListener('change', recompute);
  document.querySelectorAll('input[name=sizeMode]').forEach(function (r) {
    if (r.value === load(LS_SIZE, 'letter')) r.checked = true;
    r.addEventListener('change', function () { state.priceOv = {}; recompute(); });
  });
  $('#go').addEventListener('click', exportXlsx);
  $('#exportMap').addEventListener('click', function () {
    readMapping();
    var blob = new Blob([JSON.stringify({ mapping: state.mapping, header: load(LS_HDR, {}) }, null, 2)], { type: 'application/json' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'bang-ma-customs-code.json'; a.click();
  });
  $('#importMap').addEventListener('change', async function (e) {
    var f = e.target.files[0]; if (!f) return;
    try {
      var j = JSON.parse(await f.text());
      Object.assign(state.mapping, j.mapping || j);
      save(LS_MAP, state.mapping);
      if (j.header) save(LS_HDR, Object.assign(load(LS_HDR, {}), j.header));
      if (state.data) { renderMapping(); recompute(); }
      alert('Đã nạp bảng mã.');
    } catch (err) { alert('File không hợp lệ: ' + err.message); }
    e.target.value = '';
  });
})();
