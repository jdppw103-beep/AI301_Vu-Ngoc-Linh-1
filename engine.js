/* DataLens AI - Processing Engine (chạy 100% trên trình duyệt) */
(function (global) {
  'use strict';

  const norm = s => String(s ?? '').replace(/^\uFEFF/, '').trim().toLowerCase();

  // ---------------- Nhận diện cột ----------------
  const COLUMN_RULES = [
    { key: 'orderCode',      label: 'Mã đơn',      required: true,  kws: ['ordercode', 'order_code', 'mã đơn', 'mã đơn hàng', 'order id', 'orderid', 'so don', 'order', 'mã hđ', 'invoice', 'code'] },
    { key: 'orderDate',      label: 'Ngày',        required: false, kws: ['orderdate', 'order_date', 'ngày đơn', 'ngày', 'date', 'thời gian'] },
    { key: 'customerName',   label: 'Khách hàng',  required: false, kws: ['customername', 'khách hàng', 'tên khách', 'customer'] },
    { key: 'storeName',      label: 'Chi nhánh',   required: false, kws: ['storename', 'chi nhánh', 'cửa hàng', 'cơ sở', 'store', 'branch'] },
    { key: 'sku',            label: 'SKU',         required: false, kws: ['sku', 'mã sp', 'mã sản phẩm', 'productcode', 'product_code', 'mã hàng'] },
    { key: 'productName',    label: 'Sản phẩm',    required: true,  kws: ['productname', 'tên sản phẩm', 'tên hàng', 'sản phẩm', 'product', 'item', 'món'] },
    { key: 'quantity',       label: 'Số lượng',    required: true,  kws: ['quantity', 'số lượng', 'qty', 'sl'] },
    { key: 'unitPrice',      label: 'Đơn giá',     required: true,  kws: ['unitprice', 'unit_price', 'đơn giá', 'giá bán', 'price'] },
    { key: 'discountAmount', label: 'Chiết khấu',  required: false, kws: ['discountamount', 'discount', 'giảm giá', 'chiết khấu', 'khuyến mãi'] },
    { key: 'unitCost',       label: 'Giá vốn',     required: false, kws: ['unitcost', 'unit_cost', 'giá vốn', 'cogs', 'cost'] },
    { key: 'status',         label: 'Trạng thái',  required: false, kws: ['status', 'trạng thái', 'tình trạng'] },
    { key: 'channel',        label: 'Kênh',        required: false, kws: ['channel', 'kênh bán', 'kênh', 'nguồn'] }
  ];

  function detectColumns(headers) {
    const h = headers.map(norm);
    const used = new Set();
    const map = {};
    for (const rule of COLUMN_RULES) {
      let idx = -1;
      for (const kw of rule.kws) { const i = h.findIndex((x, j) => !used.has(j) && x === kw); if (i >= 0) { idx = i; break; } }
      if (idx < 0) for (const kw of rule.kws) { const i = h.findIndex((x, j) => !used.has(j) && x.includes(kw)); if (i >= 0) { idx = i; break; } }
      if (idx >= 0) used.add(idx);
      map[rule.key] = idx;
    }
    return map;
  }

  // ---------------- Chuẩn hoá giá trị ----------------
  function parseNumber(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    let s = String(v ?? '').trim().replace(/[^\d.,\-]/g, '');
    if (!s) return 0;
    const lastDot = s.lastIndexOf('.'), lastComma = s.lastIndexOf(',');
    if (lastDot >= 0 && lastComma >= 0) {
      s = lastComma > lastDot ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (lastComma >= 0) {
      const p = s.split(',');
      s = (p.length > 2 || p[p.length - 1].length === 3) ? s.replace(/,/g, '') : s.replace(',', '.');
    } else if (lastDot >= 0) {
      const p = s.split('.');
      if (p.length > 2 || p[p.length - 1].length === 3) s = s.replace(/\./g, '');
    }
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function parseDate(v) {
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)); // Excel serial
    const s = String(v ?? '').trim();
    if (!s) return null;
    let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  function normalizeStatus(v) {
    const s = norm(v);
    if (!s) return 'COMPLETED';
    if (/hoàn tất|hoàn thành|thành công|complete|success|done|delivered|đã giao/.test(s)) return 'COMPLETED';
    if (/cancel|hủy|huỷ/.test(s)) return 'CANCELLED';
    if (/fail|lỗi|thất bại|error/.test(s)) return 'FAILED';
    if (/return|refund|hoàn|trả/.test(s)) return 'RETURNED';
    return 'COMPLETED';
  }

  // ---------------- Đọc tệp ----------------
  function parseCSV(text) {
    text = text.replace(/^\uFEFF/, '');
    const firstLine = text.split(/\r?\n/)[0] || '';
    const counts = { ',': (firstLine.match(/,/g) || []).length, ';': (firstLine.match(/;/g) || []).length, '\t': (firstLine.match(/\t/g) || []).length };
    const delim = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
    const rows = []; let row = []; let cur = ''; let q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
        else cur += c;
      } else if (c === '"') q = true;
      else if (c === delim) { row.push(cur); cur = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(cur); cur = '';
        if (row.some(x => x.trim() !== '')) rows.push(row);
        row = [];
      } else cur += c;
    }
    row.push(cur);
    if (row.some(x => x.trim() !== '')) rows.push(row);
    return rows;
  }

  function rowsToRecords(rows) {
    if (!rows.length) throw new Error('Tệp không có dữ liệu.');
    const headers = rows[0].map(x => String(x ?? ''));
    const map = detectColumns(headers);
    const get = (r, k) => map[k] >= 0 ? r[map[k]] : undefined;
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every(x => String(x ?? '').trim() === '')) continue;
      const store = String(get(r, 'storeName') ?? '').trim() || 'Chưa phân loại';
      const qty = get(r, 'quantity');
      records.push({
        orderCode: String(get(r, 'orderCode') ?? '').trim() || `ROW-${i + 1}`,
        orderDate: parseDate(get(r, 'orderDate')),
        customerName: String(get(r, 'customerName') ?? '').trim(),
        storeName: store,
        sku: (String(get(r, 'sku') ?? '').trim() || 'N/A').toUpperCase(),
        productName: String(get(r, 'productName') ?? '').trim() || 'Không rõ tên',
        quantity: qty === undefined || String(qty).trim() === '' ? 1 : parseNumber(qty),
        unitPrice: parseNumber(get(r, 'unitPrice')),
        discountAmount: parseNumber(get(r, 'discountAmount')),
        unitCost: parseNumber(get(r, 'unitCost')),
        status: normalizeStatus(get(r, 'status')),
        channel: String(get(r, 'channel') ?? '').trim() || 'Khác'
      });
    }
    if (!records.length) throw new Error('Không nhận diện được dòng dữ liệu hợp lệ nào.');
    const mapping = COLUMN_RULES.map(c => ({ key: c.key, label: c.label, required: c.required, column: map[c.key] >= 0 ? headers[map[c.key]] : null }));
    return { records, mapping, headers };
  }

  async function parseFile(file) {
    const name = file.name.toLowerCase();
    if (name.endsWith('.json')) {
      const data = JSON.parse(await file.text());
      const arr = Array.isArray(data) ? data : (Array.isArray(data.data) ? data.data : Object.values(data).find(Array.isArray));
      if (!arr || !arr.length) throw new Error('JSON phải là một mảng các bản ghi.');
      const headers = [...new Set(arr.flatMap(o => Object.keys(o)))];
      return rowsToRecords([headers, ...arr.map(o => headers.map(h => o[h]))]);
    }
    if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
      if (!global.XLSX) throw new Error('Thư viện đọc Excel chưa tải xong, vui lòng thử lại.');
      const wb = global.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = global.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
      return rowsToRecords(rows);
    }
    return rowsToRecords(parseCSV(await file.text()));
  }

  // ---------------- Tổng quan dữ liệu thô ----------------
  const lineValue = r => r.quantity * r.unitPrice - r.discountAmount;
  const keyOf = r => `${r.orderCode}__${r.sku}__${r.productName}`.toUpperCase();

  function summarize(raw) {
    const groups = {};
    raw.forEach(r => { const k = keyOf(r); groups[k] = (groups[k] || 0) + 1; });
    return {
      totalRecords: raw.length,
      uniqueStores: new Set(raw.map(r => r.storeName)).size,
      uniqueSkus: new Set(raw.map(r => r.sku + r.productName)).size,
      duplicateGroups: Object.values(groups).filter(n => n > 1).length,
      cancelledOrders: raw.filter(r => r.status === 'CANCELLED' || r.status === 'FAILED').length,
      returnedOrders: raw.filter(r => r.status === 'RETURNED').length,
      missingCost: raw.filter(r => !r.unitCost).length,
      rawGrossValue: raw.reduce((s, r) => s + lineValue(r), 0)
    };
  }

  // ---------------- Làm sạch (Không bịa giá vốn) ----------------
  function clean(raw, cfg) {
    cfg = Object.assign({ removeDuplicates: true, excludeCancelled: true, deductReturns: true, normalizeFormat: true }, cfg || {});
    const logs = [], cleaned = [], seen = new Set();
    let dup = 0, cancel = 0, ret = 0, missingCost = 0, invalid = 0, prevented = 0;
    logs.push(`[KHỞI ĐỘNG] Quét ${raw.length} bản ghi thô…`);

    for (const r of raw) {
      const val = lineValue(r);
      if (r.quantity <= 0 || r.unitPrice <= 0 || !r.orderCode || r.productName === 'Không rõ tên') {
        invalid++; logs.push(`[DỮ LIỆU LỖI/TRỐNG] Đơn ${r.orderCode}: thiếu thông tin hoặc số lượng/đơn giá <= 0 → tách ngoại lệ.`); continue;
      }
      const k = keyOf(r);
      if (cfg.removeDuplicates && seen.has(k)) {
        dup++; prevented += val; logs.push(`[TRÙNG LẶP] Đơn ${r.orderCode} · ${r.productName} bị ghi trùng → tách riêng (${fmt(val)} đ).`); continue;
      }
      seen.add(k);
      if (cfg.excludeCancelled && (r.status === 'CANCELLED' || r.status === 'FAILED')) {
        cancel++; prevented += val; logs.push(`[LOẠI ĐƠN HỦY] Đơn ${r.orderCode} (${r.status}) → không tính doanh thu (${fmt(val)} đ).`); continue;
      }
      const unitCost = r.unitCost > 0 ? r.unitCost : 0;
      if (unitCost <= 0) missingCost++;
      let remark = unitCost <= 0 ? 'Thiếu giá vốn thực' : 'Hợp lệ';
      let gross = r.quantity * r.unitPrice, net = gross - r.discountAmount, cost = unitCost > 0 ? r.quantity * unitCost : 0;
      if (r.status === 'RETURNED') {
        ret++; remark = 'Hàng hoàn (giảm trừ)';
        if (cfg.deductReturns) {
          gross = -gross; net = -net; cost = -cost;
          logs.push(`[HÀNG HOÀN] Đơn ${r.orderCode} → giảm trừ doanh thu ${fmt(Math.abs(net))} đ.`);
        }
      }
      const profit = unitCost > 0 ? (net - cost) : 0;
      cleaned.push(Object.assign({}, r, {
        unitCost, grossRevenue: gross, netRevenue: net, totalCost: cost, grossProfit: profit,
        marginPercent: (unitCost > 0 && net > 0) ? round2(profit / net * 100) : 0, remark
      }));
    }
    const gmv = cleaned.filter(x => x.status !== 'RETURNED').reduce((s, x) => s + x.grossRevenue, 0);
    const actualRevenue = cleaned.reduce((s, x) => s + x.netRevenue, 0);
    if (missingCost) logs.push(`[LƯU Ý GIÁ VỐN] ${missingCost} dòng chưa có giá vốn thực tế (giữ nguyên 0 đ, không tự suy diễn).`);
    logs.push(`[HOÀN TẤT] ${raw.length} dòng thô → ${cleaned.length} dòng sạch. Doanh thu thực tế: ${fmt(actualRevenue)} đ (GMV: ${fmt(gmv)} đ).`);
    logs.push(`[THỐNG KÊ] Trùng: ${dup} · Hủy/lỗi: ${cancel} · Hoàn: ${ret} · Trống/sai ĐD: ${invalid} · Sai lệch ngăn chặn: ${fmt(prevented)} đ.`);
    return { cleaned, logs, stats: { rawCount: raw.length, cleanedCount: cleaned.length, duplicates: dup, cancelled: cancel, returns: ret, imputed: missingCost, invalid, prevented, gmv, actualRevenue } };
  }

  // ---------------- Phân tích ----------------
  function analyze(cleaned, stats, context) {
    context = context || {};
    const sum = (arr, f) => arr.reduce((s, x) => s + f(x), 0);
    const totalGross = sum(cleaned, x => Math.max(0, x.grossRevenue));
    const totalDiscount = sum(cleaned, x => x.status === 'RETURNED' ? 0 : x.discountAmount);
    const totalNet = sum(cleaned, x => x.netRevenue);
    const totalCost = sum(cleaned, x => x.totalCost);
    const totalProfit = totalNet - totalCost;
    const margin = totalNet > 0 ? round2(totalProfit / totalNet * 100) : 0;
    const returnOrders = cleaned.filter(x => x.status === 'RETURNED').length;
    const returnRate = cleaned.length ? round2(returnOrders / cleaned.length * 100) : 0;
    const cancelRate = stats && stats.rawCount ? round2(stats.cancelled / stats.rawCount * 100) : 0;

    const groupBy = (arr, f) => arr.reduce((m, x) => { const k = f(x); (m[k] = m[k] || []).push(x); return m; }, {});

    const stores = Object.entries(groupBy(cleaned, x => x.storeName)).map(([name, g]) => {
      const net = sum(g, x => x.netRevenue), cost = sum(g, x => x.totalCost), profit = net - cost;
      const m = net > 0 ? round2(profit / net * 100) : 0;
      const rets = g.filter(x => x.status === 'RETURNED').length;
      const rr = round2(rets / g.length * 100);
      let st = 'Vận hành tốt';
      if (m < 12) st = 'Biên lợi nhuận mỏng';
      else if (rr > 15) st = 'Đổi trả cao';
      return { storeName: name, orders: g.length, netRevenue: net, grossProfit: profit, marginPercent: m, returnRate: rr, status: st };
    }).sort((a, b) => b.netRevenue - a.netRevenue);

    const products = Object.entries(groupBy(cleaned, x => x.sku + '||' + x.productName)).map(([k, g]) => {
      const net = sum(g, x => x.netRevenue), cost = sum(g, x => x.totalCost), profit = net - cost;
      return { sku: k.split('||')[0], productName: k.split('||')[1], quantity: sum(g, x => x.status === 'RETURNED' ? -x.quantity : x.quantity), netRevenue: net, grossProfit: profit, marginPercent: net > 0 ? round2(profit / net * 100) : 0 };
    }).sort((a, b) => b.netRevenue - a.netRevenue);

    const channels = Object.entries(groupBy(cleaned, x => x.channel)).map(([name, g]) => {
      const net = sum(g, x => x.netRevenue);
      return { channel: name, orders: g.length, netRevenue: net, share: totalNet > 0 ? round2(net / totalNet * 100) : 0 };
    }).sort((a, b) => b.netRevenue - a.netRevenue);

    const customers = Object.entries(groupBy(cleaned.filter(x => x.customerName), x => x.customerName))
      .map(([name, g]) => ({ name, netRevenue: sum(g, x => x.netRevenue), orders: g.length }))
      .sort((a, b) => b.netRevenue - a.netRevenue);

    // Xu hướng theo ngày
    const dated = cleaned.filter(x => x.orderDate);
    const trend = Object.entries(groupBy(dated, x => isoDay(x.orderDate)))
      .map(([d, g]) => ({ date: d, netRevenue: sum(g, x => x.netRevenue) }))
      .sort((a, b) => a.date.localeCompare(b.date));

    // ----- Khuyến nghị -----
    const problem = norm(context.problem);
    const focus = [], recs = [];
    const lowStores = stores.filter(s => s.marginPercent < 12);
    const topCh = channels[0], topP = products[0];

    if (/hủy|huỷ|cancel/.test(problem)) focus.push(`[TRỌNG TÂM · HỦY ĐƠN] Tỷ lệ đơn hủy/lỗi là ${cancelRate}% (${stats.cancelled} đơn). ${cancelRate > 5 ? 'Vượt ngưỡng 5% — cần rà soát khâu xác nhận đơn, cổng thanh toán và thời gian giao.' : 'Đang trong ngưỡng an toàn (< 5%).'}`);
    if (/biên|lợi nhuận|chi phí|hao hụt|giá vốn|margin/.test(problem)) focus.push(`[TRỌNG TÂM · LỢI NHUẬN] Biên lợi nhuận gộp toàn hệ thống ${margin}%. ${lowStores.length ? `Các điểm bán dưới 12%: ${lowStores.map(s => s.storeName).join(', ')} — kiểm soát chiết khấu và giá vốn đầu vào.` : 'Không có điểm bán nào dưới ngưỡng 12%.'}`);
    if (/kênh|channel|phân phối/.test(problem) && topCh) focus.push(`[TRỌNG TÂM · KÊNH] '${topCh.channel}' dẫn đầu với ${topCh.share}% doanh thu thuần. ${channels.length > 1 ? `Kênh yếu nhất: '${channels[channels.length - 1].channel}' (${channels[channels.length - 1].share}%).` : ''}`);
    if (/trả|hoàn|return/.test(problem)) focus.push(`[TRỌNG TÂM · ĐỔI TRẢ] Tỷ lệ hoàn ${returnRate}% (${returnOrders} đơn). ${returnRate > 8 ? 'Cao — kiểm tra chất lượng hàng và quy trình đóng gói.' : 'Trong ngưỡng chấp nhận.'}`);
    if (/khách|churn|rời bỏ|customer/.test(problem) && customers.length) {
      const top5 = sum(customers.slice(0, 5), c => c.netRevenue);
      focus.push(`[TRỌNG TÂM · KHÁCH HÀNG] ${customers.length} khách hàng có giao dịch; top 5 chiếm ${totalNet > 0 ? round2(top5 / totalNet * 100) : 0}% doanh thu — ${top5 / Math.max(1, totalNet) > 0.5 ? 'phụ thuộc cao vào nhóm nhỏ, cần chương trình giữ chân.' : 'cơ cấu khách hàng phân tán tốt.'}`);
    }
    if (/trùng|sai lệch|doanh thu ảo/.test(problem)) focus.push(`[TRỌNG TÂM · CHẤT LƯỢNG DỮ LIỆU] Đã loại ${stats.duplicates} dòng trùng và ${stats.cancelled} đơn hủy, ngăn chặn ${fmt(stats.prevented)} đ doanh thu ảo.`);

    if (returnRate > 8) recs.push(`[VẬN HÀNH] Tỷ lệ hoàn hàng ${returnRate}% ở mức cao. Rà soát kiểm hàng trước giao và chính sách đổi trả.`);
    if (lowStores.length) recs.push(`[GIÁ & CHIẾT KHẤU] ${lowStores.length} điểm bán có biên lợi nhuận dưới 12%. Thiết lập trần chiết khấu theo cấp phê duyệt.`);
    if (cancelRate > 5) recs.push(`[ĐƠN HÀNG] Tỷ lệ hủy/lỗi ${cancelRate}% — cần quy trình xác nhận đơn tự động.`);
    if (topCh && channels.length > 1) recs.push(`[TĂNG TRƯỞNG] Kênh '${topCh.channel}' đóng góp ${topCh.share}% doanh thu. Ưu tiên ngân sách và tồn kho cho kênh này.`);
    if (topP) recs.push(`[SẢN PHẨM] '${topP.productName}' dẫn đầu doanh thu (${fmt(topP.netRevenue)} đ, biên ${topP.marginPercent}%). Đảm bảo tồn kho khả dụng.`);
    if (stats.imputed) recs.push(`[DỮ LIỆU] ${stats.imputed} dòng thiếu giá vốn đã được ước tính — bổ sung giá vốn thực để lợi nhuận chính xác hơn.`);
    if (!recs.length && !focus.length) recs.push('[ỔN ĐỊNH] Các chỉ số chính đều trong ngưỡng an toàn. Tiếp tục theo dõi định kỳ.');

    const biz = context.businessType ? ` cho mô hình ${context.businessType}` : '';
    const summary = `Phân tích ${cleaned.length} giao dịch hợp lệ${biz}: doanh thu thuần ${fmt(totalNet)} đ, lợi nhuận gộp ${fmt(totalProfit)} đ (biên ${margin}%). Đã loại ${stats.duplicates} dòng trùng, ${stats.cancelled} đơn hủy/lỗi và hạch toán ${stats.returns} đơn hoàn.`;

    return {
      summary,
      metrics: { totalGross, totalDiscount, totalNet, totalCost, totalProfit, margin, orders: cleaned.length, returnOrders, returnRate, cancelRate },
      stores, products, channels, customers, trend,
      recommendations: focus.concat(recs)
    };
  }

  // ---------------- Hỏi đáp ----------------
  function answer(q, a) {
    q = norm(q);
    const m = a.metrics;
    const store = a.stores.find(s => q.includes(norm(s.storeName)));
    if (store) return `${store.storeName}: doanh thu thuần ${fmt(store.netRevenue)} đ, lợi nhuận gộp ${fmt(store.grossProfit)} đ (biên ${store.marginPercent}%), ${store.orders} giao dịch, tỷ lệ hoàn ${store.returnRate}%. Đánh giá: ${store.status}.`;
    const prod = a.products.find(p => q.includes(norm(p.productName)) || (p.sku !== 'N/A' && q.includes(norm(p.sku))));
    if (prod) return `${prod.productName}: bán ${prod.quantity} sản phẩm, doanh thu thuần ${fmt(prod.netRevenue)} đ, lợi nhuận ${fmt(prod.grossProfit)} đ (biên ${prod.marginPercent}%).`;
    if (/chi nhánh|cửa hàng|điểm bán|rủi ro/.test(q)) {
      const best = a.stores[0], worst = [...a.stores].sort((x, y) => x.marginPercent - y.marginPercent)[0];
      return best ? `Dẫn đầu doanh thu: ${best.storeName} (${fmt(best.netRevenue)} đ). Biên lợi nhuận thấp nhất: ${worst.storeName} (${worst.marginPercent}%, hoàn ${worst.returnRate}%) — nên ưu tiên hỗ trợ.` : 'Dữ liệu chưa có thông tin chi nhánh.';
    }
    if (/sản phẩm|bán chạy|mặt hàng|sku/.test(q)) {
      const top = a.products.slice(0, 3).map((p, i) => `${i + 1}. ${p.productName} (${fmt(p.netRevenue)} đ)`).join('; ');
      return `Top sản phẩm theo doanh thu: ${top}.`;
    }
    if (/kênh|channel/.test(q)) return `Cơ cấu kênh: ${a.channels.map(c => `${c.channel} ${c.share}%`).join(', ')}.`;
    if (/khách/.test(q)) return a.customers.length ? `Top khách hàng: ${a.customers.slice(0, 3).map(c => `${c.name} (${fmt(c.netRevenue)} đ)`).join('; ')}.` : 'Dữ liệu không có cột khách hàng.';
    if (/hoàn|trả/.test(q)) return `Tỷ lệ hoàn ${m.returnRate}% (${m.returnOrders} đơn). Các đơn hoàn đã được giảm trừ doanh thu và hoàn nhập giá vốn.`;
    if (/hủy|huỷ/.test(q)) return `Tỷ lệ đơn hủy/lỗi trên dữ liệu thô là ${m.cancelRate}%. Các đơn này đã bị loại khỏi doanh thu.`;
    if (/doanh thu|lợi nhuận|biên|tổng/.test(q)) return `Doanh thu thuần ${fmt(m.totalNet)} đ, giá vốn ${fmt(m.totalCost)} đ, lợi nhuận gộp ${fmt(m.totalProfit)} đ (biên ${m.margin}%).`;
    return `${a.summary} Bạn có thể hỏi về chi nhánh, sản phẩm, kênh bán, khách hàng, đổi trả hoặc hủy đơn.`;
  }

  // ---------------- Xuất ----------------
  function toCSV(cleaned) {
    const cols = ['Mã đơn', 'Ngày', 'Khách hàng', 'Chi nhánh', 'SKU', 'Sản phẩm', 'Số lượng', 'Đơn giá', 'Chiết khấu', 'Giá vốn ĐV', 'Doanh thu thuần', 'Tổng giá vốn', 'Lợi nhuận gộp', 'Biên LN %', 'Trạng thái', 'Kênh', 'Ghi chú'];
    const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const lines = cleaned.map(r => [q(r.orderCode), q(r.orderDate ? isoDay(r.orderDate) : ''), q(r.customerName), q(r.storeName), q(r.sku), q(r.productName),
      r.quantity, r.unitPrice, r.discountAmount, Math.round(r.unitCost), Math.round(r.netRevenue), Math.round(r.totalCost), Math.round(r.grossProfit), r.marginPercent, q(r.status), q(r.channel), q(r.remark)].join(','));
    return '\uFEFF' + cols.join(',') + '\n' + lines.join('\n');
  }

  function templateCSV() {
    return '\uFEFFMã đơn,Ngày,Khách hàng,Chi nhánh,SKU,Sản phẩm,Số lượng,Đơn giá,Chiết khấu,Giá vốn,Trạng thái,Kênh\n';
  }

  // ---------------- Tiện ích ----------------
  function round2(n) { return Math.round(n * 100) / 100; }
  function fmt(n) { return Math.round(n || 0).toLocaleString('vi-VN'); }
  function isoDay(d) { const z = n => String(n).padStart(2, '0'); return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`; }

  global.DataLens = { COLUMN_RULES, parseFile, parseCSV, rowsToRecords, summarize, clean, analyze, answer, toCSV, templateCSV, parseNumber, normalizeStatus, fmt };
})(window);

