using System.Text;
using System.Text.Json;

public static class HtmlDashboardExporter
{
    public static string BuildStandaloneHtml(DataAnalystHub hub, DeepAnalysisResult analysis)
    {
        var m = analysis.OverallMetrics;
        var sch = analysis.Schema;
        var reportDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

        // Nhúng dữ liệu tổng hợp theo chiều phân tích (tuyệt đối KHÔNG nhúng tên khách hàng / PII)
        var compactRows = hub.CleanedData
            .GroupBy(r => new
            {
                d = r.OrderDate.ToString("yyyy-MM-dd"),
                s = r.StoreName,
                p = r.ProductName,
                c = r.Channel,
                st = r.Status
            })
            .Select(g => new
            {
                d = g.Key.d,
                s = g.Key.s,
                p = g.Key.p,
                c = g.Key.c,
                st = g.Key.st,
                n = g.Count(),
                q = g.Sum(x => x.Quantity),
                gmv = g.Sum(x => x.GrossRevenue),
                disc = g.Sum(x => x.DiscountAmount),
                net = g.Sum(x => x.NetRevenue),
                cost = g.Sum(x => x.TotalCost),
                prof = g.Sum(x => x.GrossProfit)
            })
            .OrderBy(x => x.d)
            .ToList();

        var payload = new
        {
            meta = new
            {
                brand = "Ngọc Linh Data Studio",
                reportTitle = "EXECUTIVE BUSINESS INTELLIGENCE DASHBOARD",
                sourceName = hub.SourceName,
                generatedAt = reportDate,
                processedAt = hub.ProcessedAt,
                dataVersion = hub.DataVersion,
                isVerified = hub.IsHumanVerified,
                rawCount = hub.LastStats.RawCount,
                cleanedCount = hub.CleanedData.Count,
                cancelledExcluded = hub.LastStats.CancelledExcluded,
                duplicatesRemoved = hub.LastStats.DuplicatesRemoved,
                returnsAdjusted = hub.LastStats.ReturnsAdjusted,
                cancelRatePercent = m.CancelRatePercent,
                preventedOvercount = m.PreventedOvercountValue,
                growthRatePercent = analysis.GrowthRatePercent,
                totalCustomers = analysis.TotalCustomers,
                repeatCustomers = analysis.RepeatCustomers
            },
            schema = new
            {
                hasDate = sch.HasDate,
                hasRevenue = sch.HasRevenue,
                hasCost = m.HasCostData,
                hasProduct = sch.HasProduct,
                hasStore = sch.HasStore,
                hasCustomer = sch.HasCustomer,
                hasChannel = sch.HasChannel,
                hasDiscount = sch.HasDiscount
            },
            rows = compactRows,
            customers = analysis.CustomerPerformance.Take(10).Select(c => new { name = c.CustomerName, orders = c.OrderCount, units = c.TotalUnits, net = c.NetRevenue }).ToList(),
            insights = analysis.StructuredRecommendations.Select(r => new
            {
                category = r.Category,
                priority = r.Priority,
                title = r.Title,
                content = r.Content
            }).ToList(),
            anomalies = analysis.Anomalies.Select(a => new
            {
                type = a.Type,
                severity = a.Severity,
                title = a.Dimension,
                evidence = a.MetricValue,
                cause = a.RootCause
            }).ToList()
        };

        var jsonPayload = JsonSerializer.Serialize(payload);

        var sb = new StringBuilder();
        sb.Append("""
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ngọc Linh Data Studio · Executive BI Dashboard</title>
<style>
  :root {
    --bg: #FCF8F2;
    --surface: #FFFFFF;
    --soft: #F5EFE6;
    --ink: #0E0907;
    --brown: #664F41;
    --brown-soft: #90735F;
    --beige: #C9A991;
    --amber: #DD9E5E;
    --blue: #657DC4;
    --ok: #3D7A52;
    --red: #C84B4B;
    --border: rgba(102, 79, 65, 0.14);
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif; background: var(--bg); color: var(--ink); line-height: 1.5; }
  .layout { display: flex; min-height: 100vh; }
  .nav { width: 230px; flex-shrink: 0; background: var(--soft); border-right: 1px solid var(--border); padding: 20px 14px; position: sticky; top: 0; height: 100vh; overflow-y: auto; }
  .nav-brand { font-size: 16px; font-weight: 800; color: var(--ink); }
  .nav-sub { font-size: 11px; color: var(--brown-soft); margin-bottom: 18px; }
  .nav a { display: block; padding: 9px 12px; border-radius: 10px; text-decoration: none; color: var(--brown); font-size: 12.5px; font-weight: 600; margin-bottom: 4px; }
  .nav a:hover, .nav a.active { background: var(--surface); color: var(--ink); box-shadow: 0 4px 12px -6px rgba(14,9,7,0.12); }
  .main { flex: 1; padding: 22px 26px; max-width: 1380px; min-width: 0; }
  .header { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 14px; padding: 18px 22px; background: var(--surface); border-radius: 20px; border: 1px solid var(--border); margin-bottom: 18px; }
  .h-title { font-size: 19px; font-weight: 800; letter-spacing: -0.02em; }
  .h-sub { font-size: 12px; color: var(--brown-soft); margin-top: 2px; }
  .badges { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
  .badge { padding: 5px 12px; border-radius: 999px; font-size: 11.5px; font-weight: 700; background: var(--soft); color: var(--brown); }
  .badge.ok { background: #E9F4ED; color: var(--ok); }
  .badge.warn { background: #FDF2E4; color: #9A5B1A; }
  .btn { border: none; cursor: pointer; border-radius: 999px; padding: 8px 16px; font-size: 12px; font-weight: 700; }
  .btn-primary { background: var(--brown); color: #fff; }
  .btn-ghost { background: var(--soft); color: var(--brown); }
  .filters { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; background: var(--surface); padding: 14px 18px; border-radius: 18px; border: 1px solid var(--border); margin-bottom: 18px; align-items: end; }
  .field label { display: block; font-size: 10.5px; font-weight: 700; text-transform: uppercase; color: var(--brown-soft); margin-bottom: 4px; }
  .field input, .field select { width: 100%; padding: 7px 10px; border-radius: 10px; border: 1px solid var(--border); background: var(--bg); color: var(--ink); font-size: 12.5px; }
  .sec-head { font-size: 15px; font-weight: 800; margin: 22px 0 10px; color: var(--ink); display: flex; align-items: center; gap: 8px; }
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin-bottom: 16px; }
  .card { background: var(--surface); border-radius: 20px; padding: 18px; border: 1px solid var(--border); box-shadow: 0 8px 24px -16px rgba(14,9,7,0.08); }
  .card-hero { background: linear-gradient(135deg, #664F41 0%, #45342A 100%); color: #FCF8F2; border: none; }
  .kpi-lbl { font-size: 11px; font-weight: 700; text-transform: uppercase; color: var(--brown-soft); }
  .card-hero .kpi-lbl { color: #E8D8CA; }
  .kpi-val { font-size: 22px; font-weight: 800; margin-top: 4px; font-variant-numeric: tabular-nums; }
  .kpi-note { font-size: 11px; color: var(--brown-soft); margin-top: 4px; }
  .card-hero .kpi-note { color: #DEC8B6; }
  .grid-2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap: 14px; margin-bottom: 16px; }
  .c-title { font-size: 14px; font-weight: 800; margin: 0 0 2px; }
  .c-sub { font-size: 11.5px; color: var(--brown-soft); margin: 0 0 12px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th { text-align: left; font-size: 10.5px; text-transform: uppercase; color: var(--brown-soft); padding: 8px; border-bottom: 1px solid var(--border); cursor: pointer; }
  td { padding: 8px; border-bottom: 1px solid rgba(102,79,65,0.07); }
  .num { text-align: right; font-variant-numeric: tabular-nums; }
  .alert { padding: 12px 14px; border-radius: 14px; margin-bottom: 8px; border-left: 4px solid var(--brown); background: var(--bg); font-size: 12.5px; }
  .alert.critical { border-left-color: var(--red); background: #FDF3F3; }
  .alert.warning { border-left-color: var(--amber); background: #FDF7EE; }
  .alert.positive { border-left-color: var(--ok); background: #EFF7F1; }
  .missing-box { padding: 24px; text-align: center; background: var(--bg); border-radius: 14px; color: var(--brown-soft); font-size: 12.5px; border: 1px dashed var(--border); }
  .notice { margin-top: 20px; padding: 14px 18px; border-radius: 16px; background: var(--soft); font-size: 11.5px; color: var(--brown); display: flex; flex-wrap: wrap; justify-content: space-between; gap: 10px; }
  @media (max-width: 900px) { .nav { display: none; } .grid-2 { grid-template-columns: 1fr; } }
  @media print { .nav, .filters, .no-print { display: none !important; } .main { padding: 0; } .card { break-inside: avoid; box-shadow: none; } }
</style>
</head>
<body>
<div class="layout">
  <aside class="nav no-print">
    <div class="nav-brand">Ngọc Linh Data Studio</div>
    <div class="nav-sub">Executive BI Dashboard</div>
    <a href="#sec-kpi" class="active">1. Tổng quan &amp; KPI</a>
    <a href="#sec-fin">2. Tài chính &amp; Xu hướng</a>
    <a href="#sec-prod">3. Phân tích Sản phẩm</a>
    <a href="#sec-store">4. Chi nhánh &amp; Kênh bán</a>
    <a href="#sec-ops">5. Vận hành &amp; Chiết khấu</a>
    <a href="#sec-cust">6. Khách hàng</a>
    <a href="#sec-ins">7. Insight &amp; Cảnh báo</a>
    <a href="#sec-exp">8. Data Explorer</a>
  </aside>

  <main class="main">
    <div class="header">
      <div>
        <div class="h-title">EXECUTIVE BUSINESS INTELLIGENCE DASHBOARD</div>
        <div class="h-sub" id="metaLine">Ngọc Linh Data Studio</div>
      </div>
      <div class="badges">
        <span class="badge" id="verBadge">Phiên bản #1</span>
        <span class="badge" id="verifyBadge">Kiểm duyệt</span>
        <button class="btn btn-primary no-print" onclick="window.print()">In / Lưu PDF</button>
      </div>
    </div>

    <div class="filters no-print">
      <div class="field"><label>Từ ngày</label><input type="date" id="fFrom"></div>
      <div class="field"><label>Đến ngày</label><input type="date" id="fTo"></div>
      <div class="field"><label>Chi nhánh</label><select id="fStore"><option value="ALL">Tất cả chi nhánh</option></select></div>
      <div class="field"><label>Sản phẩm</label><select id="fProd"><option value="ALL">Tất cả sản phẩm</option></select></div>
      <div class="field"><label>Kênh bán</label><select id="fChan"><option value="ALL">Tất cả kênh</option></select></div>
      <div class="field"><label>Trạng thái</label><select id="fStatus"><option value="ALL">Tất cả trạng thái</option><option value="COMPLETED">COMPLETED</option><option value="RETURNED">RETURNED</option></select></div>
      <div class="field"><button class="btn btn-ghost" style="width:100%" onclick="resetFilters()">Đặt lại</button></div>
    </div>

    <div id="sec-kpi" class="sec-head">1. Hệ thống KPI điều hành</div>
    <div class="kpis" id="kpiGrid"></div>

    <div id="sec-fin" class="sec-head">2. Phân tích Tài chính &amp; Xu hướng</div>
    <div class="grid-2">
      <div class="card">
        <h3 class="c-title">Xu hướng Doanh thu thuần theo thời gian</h3>
        <p class="c-sub">Line Chart · Đơn vị: VNĐ (cập nhật theo bộ lọc)</p>
        <div id="chartLine"></div>
      </div>
      <div class="card">
        <h3 class="c-title">So sánh Doanh thu gộp (GMV) &amp; Doanh thu thuần</h3>
        <p class="c-sub">Grouped Bar Chart theo ngày giao dịch</p>
        <div id="chartGmvNet"></div>
      </div>
    </div>

    <div id="sec-prod" class="sec-head">3. Phân tích Sản phẩm &amp; Pareto</div>
    <div class="grid-2">
      <div class="card">
        <h3 class="c-title">Top 10 Sản phẩm theo Doanh thu thuần</h3>
        <p class="c-sub">Horizontal Bar Chart · Xếp hạng đóng góp doanh thu</p>
        <div id="chartTopProd"></div>
      </div>
      <div class="card">
        <h3 class="c-title">Phân tích Pareto &amp; Tỷ suất đóng góp sản phẩm</h3>
        <p class="c-sub">Nhận diện nhóm SKU chủ lực đóng góp phần lớn doanh thu</p>
        <div id="tblPareto"></div>
      </div>
    </div>

    <div id="sec-store" class="sec-head">4. Chi nhánh &amp; Kênh bán hàng</div>
    <div class="grid-2">
      <div class="card">
        <h3 class="c-title">Bảng xếp hạng &amp; Hiệu quả Chi nhánh</h3>
        <p class="c-sub">So sánh doanh thu, lợi nhuận và tỷ lệ hoàn trả</p>
        <div id="tblStores"></div>
      </div>
      <div class="card">
        <h3 class="c-title">Tỷ trọng Doanh thu theo Kênh bán hàng</h3>
        <p class="c-sub">Donut Chart &amp; Bảng cơ cấu kênh</p>
        <div id="chartChannel"></div>
      </div>
    </div>

    <div id="sec-ops" class="sec-head">5. Vận hành, Đơn hàng &amp; Chiết khấu</div>
    <div class="grid-2">
      <div class="card">
        <h3 class="c-title">Cơ cấu Trạng thái Đơn hàng</h3>
        <p class="c-sub">Hoàn tất, Hoàn trả và Đơn hủy loại trừ</p>
        <div id="chartOps"></div>
      </div>
      <div class="card">
        <h3 class="c-title">Phân tích Chính sách Chiết khấu</h3>
        <p class="c-sub">So sánh doanh thu trước và sau chiết khấu theo chi nhánh</p>
        <div id="chartDiscount"></div>
      </div>
    </div>

    <div id="sec-cust" class="sec-head">6. Phân tích Khách hàng</div>
    <div class="card" id="custBox" style="margin-bottom:16px"></div>

    <div id="sec-ins" class="sec-head">7. Business Insights &amp; Cảnh báo điều hành</div>
    <div class="card" id="alertsBox" style="margin-bottom:16px"></div>

    <div id="sec-exp" class="sec-head">8. Data Explorer (Phân trang dữ liệu sạch)</div>
    <div class="card">
      <div style="display:flex;flex-wrap:wrap;justify-content:space-between;gap:10px;margin-bottom:12px">
        <input type="text" id="expSearch" placeholder="Tìm theo ngày, chi nhánh, sản phẩm, kênh..." oninput="expPage=1;renderExplorer()" style="padding:7px 12px;border-radius:10px;border:1px solid var(--border);width:280px;background:var(--bg)">
        <div style="font-size:12px;color:var(--brown-soft)" id="expInfo"></div>
      </div>
      <div style="overflow-x:auto">
        <table>
          <thead><tr>
            <th onclick="sortExp('d')">Ngày</th><th onclick="sortExp('s')">Chi nhánh</th><th onclick="sortExp('p')">Sản phẩm</th>
            <th onclick="sortExp('c')">Kênh</th><th class="num" onclick="sortExp('q')">SL</th><th class="num" onclick="sortExp('net')">DT Thuần</th>
            <th class="num" onclick="sortExp('prof')">LN Gộp</th><th>Trạng thái</th>
          </tr></thead>
          <tbody id="expBody"></tbody>
        </table>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:10px" class="no-print">
        <span style="font-size:11.5px;color:var(--brown-soft)" id="expPageTxt"></span>
        <div style="display:flex;gap:6px">
          <button class="btn btn-ghost" onclick="if(expPage>1){expPage--;renderExplorer()}">Trước</button>
          <button class="btn btn-ghost" onclick="expPage++;renderExplorer()">Sau</button>
        </div>
      </div>
    </div>

    <div class="notice">
      <div><strong>Phạm vi dữ liệu nhúng:</strong> Chỉ bao gồm dữ liệu tổng hợp theo Thời gian, Chi nhánh, Sản phẩm, Kênh bán và giá trị tài chính; <strong>không</strong> nhúng thông tin cá nhân khách hàng (PII).</div>
      <div id="footerTime"></div>
    </div>
  </main>
</div>

<script>
const DATA = 
""");
        sb.Append(jsonPayload);
        sb.Append("""
;
const fmt = n => Math.round(Number(n || 0)).toLocaleString('vi-VN');
const esc = s => String(s ?? '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
let filteredRows = [], expPage = 1, expSort = 'net', expDir = -1;

function init() {
  const m = DATA.meta;
  document.getElementById('metaLine').textContent = `Nguồn: ${m.sourceName} · Dữ liệu sạch: ${fmt(m.cleanedCount)}/${fmt(m.rawCount)} dòng · Cập nhật: ${m.generatedAt}`;
  document.getElementById('verBadge').textContent = `Phiên bản #${m.dataVersion}`;
  const vb = document.getElementById('verifyBadge');
  vb.textContent = m.isVerified ? 'Đã kiểm duyệt' : 'Chưa kiểm duyệt';
  vb.className = 'badge ' + (m.isVerified ? 'ok' : 'warn');
  document.getElementById('footerTime').textContent = `Ngọc Linh Data Studio · ${m.generatedAt}`;

  const dates = DATA.rows.map(r => r.d).filter(Boolean).sort();
  if (dates.length) { document.getElementById('fFrom').value = dates[0]; document.getElementById('fTo').value = dates[dates.length - 1]; }
  const fill = (id, vals) => {
    const sel = document.getElementById(id);
    [...new Set(vals)].sort().forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
  };
  fill('fStore', DATA.rows.map(r => r.s));
  fill('fProd', DATA.rows.map(r => r.p));
  fill('fChan', DATA.rows.map(r => r.c));

  ['fFrom','fTo','fStore','fProd','fChan','fStatus'].forEach(id => document.getElementById(id).addEventListener('change', () => { expPage = 1; renderAll(); }));
  renderCustomerSection();
  renderAlerts();
  renderAll();
}

function resetFilters() {
  const dates = DATA.rows.map(r => r.d).filter(Boolean).sort();
  document.getElementById('fFrom').value = dates[0] || '';
  document.getElementById('fTo').value = dates[dates.length - 1] || '';
  ['fStore','fProd','fChan','fStatus'].forEach(id => document.getElementById(id).value = 'ALL');
  expPage = 1;
  renderAll();
}

function renderAll() {
  const from = document.getElementById('fFrom').value, to = document.getElementById('fTo').value;
  const st = document.getElementById('fStore').value, pr = document.getElementById('fProd').value;
  const ch = document.getElementById('fChan').value, statusF = document.getElementById('fStatus').value;

  filteredRows = DATA.rows.filter(r =>
    (!from || r.d >= from) && (!to || r.d <= to) &&
    (st === 'ALL' || r.s === st) && (pr === 'ALL' || r.p === pr) &&
    (ch === 'ALL' || r.c === ch) && (statusF === 'ALL' || r.st === statusF)
  );

  let gmv = 0, disc = 0, net = 0, cost = 0, prof = 0, orders = 0, comp = 0, ret = 0, retVal = 0, qty = 0;
  const byDay = {}, byStore = {}, byProd = {}, byChan = {};
  for (const r of filteredRows) {
    if (r.st !== 'RETURNED') { gmv += r.gmv; disc += r.disc; qty += r.q; comp += r.n; }
    else { ret += r.n; retVal += Math.abs(r.net); }
    net += r.net; cost += r.cost; prof += r.prof; orders += r.n;
    if (!byDay[r.d]) byDay[r.d] = { gmv: 0, net: 0, prof: 0, orders: 0 };
    byDay[r.d].gmv += (r.st !== 'RETURNED' ? r.gmv : 0); byDay[r.d].net += r.net; byDay[r.d].prof += r.prof; byDay[r.d].orders += r.n;
    if (!byStore[r.s]) byStore[r.s] = { net: 0, prof: 0, disc: 0, gmv: 0, orders: 0, ret: 0 };
    byStore[r.s].net += r.net; byStore[r.s].prof += r.prof; byStore[r.s].disc += r.disc; byStore[r.s].gmv += r.gmv; byStore[r.s].orders += r.n; if (r.st === 'RETURNED') byStore[r.s].ret += r.n;
    if (!byProd[r.p]) byProd[r.p] = { q: 0, net: 0, prof: 0 };
    byProd[r.p].q += r.q; byProd[r.p].net += r.net; byProd[r.p].prof += r.prof;
    if (!byChan[r.c]) byChan[r.c] = { net: 0, orders: 0 };
    byChan[r.c].net += r.net; byChan[r.c].orders += r.n;
  }

  const days = Object.entries(byDay).sort((a,b) => a[0].localeCompare(b[0]));
  const hasCost = DATA.schema.hasCost && cost > 0;
  const margin = hasCost && net > 0 ? ((prof / net) * 100).toFixed(1) : '0.0';
  const costRatio = hasCost && net > 0 ? ((cost / net) * 100).toFixed(1) : '0.0';
  const discRate = gmv > 0 ? ((disc / gmv) * 100).toFixed(2) : '0.00';
  const aov = comp > 0 ? net / comp : 0;
  const compRate = orders > 0 ? ((comp / orders) * 100).toFixed(1) : '0.0';
  const retRate = orders > 0 ? ((ret / orders) * 100).toFixed(2) : '0.00';
  const growth = days.length >= 2 && days[0][1].net > 0
    ? (((days[days.length-1][1].net - days[0][1].net) / days[0][1].net) * 100).toFixed(1)
    : Number(DATA.meta.growthRatePercent || 0).toFixed(1);

  const kpis = [
    { hero: true, lbl: 'Doanh thu thuần (Net Revenue)', val: fmt(net) + ' đ', note: `Tỷ lệ giữ từ GMV: ${gmv>0?((net/gmv)*100).toFixed(1):'0.0'}%` },
    { lbl: 'Tổng doanh thu gộp (GMV)', val: fmt(gmv) + ' đ', note: `Chiết khấu: ${fmt(disc)} đ (${discRate}%)` },
    hasCost
      ? { lbl: 'Lợi nhuận gộp & Biên LN', val: fmt(prof) + ' đ', note: `Biên LN gộp: ${margin}% (Công thức: LN / DT)` }
      : { lbl: 'Lợi nhuận gộp', val: '—', note: '' },
    hasCost
      ? { lbl: 'Tổng giá vốn (COGS)', val: fmt(cost) + ' đ', note: `Tỷ trọng giá vốn / DT: ${costRatio}%` }
      : { lbl: 'Tổng giá vốn (COGS)', val: '0 đ', note: 'Chưa có dữ liệu giá vốn đầu vào' },
    { lbl: 'Giá trị đơn TB (AOV)', val: fmt(aov) + ' đ', note: `Tổng sản lượng bán: ${fmt(qty)} đơn vị` },
    { lbl: 'Đơn hợp lệ & Hoàn tất', val: `${fmt(orders)} đơn`, note: `Hoàn tất: ${fmt(comp)} (${compRate}%)` },
    { lbl: 'Tỷ lệ hoàn trả (Return)', val: `${retRate}%`, note: `${fmt(ret)} đơn hoàn · Giá trị: ${fmt(retVal)} đ` },
    { lbl: 'Tỷ lệ hủy & Ảo đã chặn', val: `${DATA.meta.cancelRatePercent}%`, note: `${fmt(DATA.meta.cancelledExcluded)} hủy · Chặn ${fmt(DATA.meta.preventedOvercount)} đ` },
    { lbl: 'Tăng trưởng kỳ (Growth)', val: `${growth}%`, note: days.length >= 2 ? `Từ ${days[0][0]} → ${days[days.length-1][0]}` : 'Dữ liệu trong 1 ngày' }
  ];
  if (DATA.schema.hasCustomer) {
    const repRate = DATA.meta.totalCustomers > 0 ? ((DATA.meta.repeatCustomers / DATA.meta.totalCustomers) * 100).toFixed(1) : '0.0';
    kpis.push({ lbl: 'Khách hàng & Quay lại', val: `${fmt(DATA.meta.totalCustomers)} KH`, note: `Quay lại: ${fmt(DATA.meta.repeatCustomers)} (${repRate}%)` });
  }
  document.getElementById('kpiGrid').innerHTML = kpis.map(k => `
    <div class="card ${k.hero ? 'card-hero' : ''}">
      <div class="kpi-lbl">${k.lbl}</div>
      <div class="kpi-val">${k.val}</div>
      <div class="kpi-note">${k.note}</div>
    </div>
  `).join('');

  renderLineChart(days);
  renderGmvNetChart(days);
  renderTopProducts(Object.entries(byProd).sort((a,b) => b[1].net - a[1].net), net, hasCost);
  renderStoresTable(Object.entries(byStore).sort((a,b) => b[1].net - a[1].net), hasCost);
  renderChannelDonut(Object.entries(byChan).sort((a,b) => b[1].net - a[1].net), net);
  renderOpsChart(comp, ret, DATA.meta.cancelledExcluded);
  renderDiscountChart(Object.entries(byStore).sort((a,b) => b[1].disc - a[1].disc));
  renderExplorer();
}

function renderLineChart(days) {
  const el = document.getElementById('chartLine');
  if (!days.length) { el.innerHTML = '<div class="missing-box">Không có dữ liệu trong khoảng lọc.</div>'; return; }
  const W = 520, H = 200, pl = 50, pr = 14, pt = 14, pb = 28;
  const max = Math.max(...days.map(d => d[1].net), 1), min = Math.min(...days.map(d => d[1].net), 0), span = Math.max(max - min, 1);
  const x = i => pl + (days.length === 1 ? (W-pl-pr)/2 : i*(W-pl-pr)/(days.length-1));
  const y = v => pt + (H-pt-pb)*(1-(v-min)/span);
  const pts = days.map((d,i) => `${x(i).toFixed(1)},${y(d[1].net).toFixed(1)}`).join(' ');
  el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" style="width:100%"><line x1="${pl}" y1="${H-pb}" x2="${W-pr}" y2="${H-pb}" stroke="rgba(102,79,65,0.2)"/><polyline points="${pts}" fill="none" stroke="#664F41" stroke-width="2.5"/><text x="${pl-5}" y="${pt+8}" font-size="9.5" text-anchor="end" fill="#90735F">${fmt(max/1000)}k</text><text x="${pl-5}" y="${H-pb}" font-size="9.5" text-anchor="end" fill="#90735F">${fmt(min/1000)}k</text></svg>`;
}

function renderGmvNetChart(days) {
  const el = document.getElementById('chartGmvNet');
  const slice = days.slice(-10);
  if (!slice.length) { el.innerHTML = '<div class="missing-box">Không có dữ liệu.</div>'; return; }
  const max = Math.max(...slice.map(d => Math.max(d[1].gmv, d[1].net)), 1);
  el.innerHTML = slice.map(([dt, v]) => `
    <div style="margin-bottom:6px;font-size:11px">
      <div style="display:flex;justify-content:space-between"><span>${dt.slice(5)}</span><span>GMV ${fmt(v.gmv)} · Thuần ${fmt(v.net)}</span></div>
      <div style="display:flex;gap:3px;height:7px;margin-top:2px">
        <div style="height:100%;width:${Math.max(2,Math.round(v.gmv/max*100))}%;background:#C9A991;border-radius:4px"></div>
        <div style="height:100%;width:${Math.max(2,Math.round(v.net/max*100))}%;background:#664F41;border-radius:4px"></div>
      </div>
    </div>
  `).join('');
}

function renderTopProducts(entries, totalNet, hasCost) {
  const el = document.getElementById('chartTopProd'), pel = document.getElementById('tblPareto');
  if (!DATA.schema.hasProduct || !entries.length) {
    const msg = '<div class="missing-box">Chưa đủ dữ liệu để thực hiện phân tích sản phẩm (cần bổ sung trường Tên sản phẩm / SKU).</div>';
    el.innerHTML = msg; pel.innerHTML = msg; return;
  }
  const top10 = entries.slice(0, 10), max = Math.max(...top10.map(e => e[1].net), 1);
  el.innerHTML = top10.map(([name, v]) => `
    <div style="margin-bottom:8px">
      <div style="display:flex;justify-content:space-between;font-size:11.5px"><strong>${esc(name)}</strong><span>${fmt(v.net)} đ (${fmt(v.q)} sp)</span></div>
      <div style="height:8px;background:var(--soft);border-radius:99px"><div style="height:100%;width:${Math.max(3,Math.round(v.net/max*100))}%;background:linear-gradient(90deg,#664F41,#DD9E5E);border-radius:99px"></div></div>
    </div>
  `).join('');
  let cum = 0;
  pel.innerHTML = `<table><thead><tr><th>Sản phẩm</th><th class="num">DT Thuần</th><th class="num">Tỷ trọng</th><th class="num">Lũy kế</th></tr></thead><tbody>` +
    top10.slice(0, 7).map(([name, v]) => {
      const share = totalNet > 0 ? (v.net / totalNet) * 100 : 0; cum += share;
      return `<tr><td>${esc(name)}</td><td class="num">${fmt(v.net)}</td><td class="num">${share.toFixed(1)}%</td><td class="num"><strong>${cum.toFixed(1)}%</strong></td></tr>`;
    }).join('') + `</tbody></table>`;
}

function renderStoresTable(entries, hasCost) {
  const el = document.getElementById('tblStores');
  if (!DATA.schema.hasStore || !entries.length) { el.innerHTML = '<div class="missing-box">Chưa đủ dữ liệu để phân tích chi nhánh (cần bổ sung trường Chi nhánh).</div>'; return; }
  el.innerHTML = `<table><thead><tr><th>Chi nhánh</th><th class="num">Đơn</th><th class="num">DT Thuần</th><th class="num">LN Gộp</th><th class="num">Hoàn %</th></tr></thead><tbody>` +
    entries.slice(0, 10).map(([name, v]) => `<tr><td><strong>${esc(name)}</strong></td><td class="num">${fmt(v.orders)}</td><td class="num">${fmt(v.net)} đ</td><td class="num">${hasCost ? fmt(v.prof) + ' đ' : '—'}</td><td class="num">${v.orders ? ((v.ret/v.orders)*100).toFixed(1) : '0.0'}%</td></tr>`).join('') +
    `</tbody></table>`;
}

function renderChannelDonut(entries, totalNet) {
  const el = document.getElementById('chartChannel');
  if (!DATA.schema.hasChannel || !entries.length) { el.innerHTML = '<div class="missing-box">Chưa đủ dữ liệu để phân tách Kênh bán hàng (cần bổ sung trường Kênh).</div>'; return; }
  const cols = ['#664F41', '#DD9E5E', '#657DC4', '#3D7A52', '#C9A991'];
  const circ = 2 * Math.PI * 44;
  let off = 0;
  const segs = entries.slice(0, 6).map(([name, v], i) => {
    const ratio = totalNet > 0 ? Math.max(0, v.net / totalNet) : 0;
    const len = ratio * circ;
    const c = cols[i % cols.length];
    const svg = `<circle cx="60" cy="60" r="44" fill="none" stroke="${c}" stroke-width="18" stroke-dasharray="${len.toFixed(1)} ${circ.toFixed(1)}" stroke-dashoffset="${(-off).toFixed(1)}" transform="rotate(-90 60 60)"/>`;
    off += len;
    return { svg, name, orders: v.orders, net: v.net, pct: (ratio * 100).toFixed(1), c };
  });
  el.innerHTML = `<div style="display:flex;flex-wrap:wrap;align-items:center;gap:18px">
    <svg viewBox="0 0 120 120" style="width:136px;height:136px;flex-shrink:0">${segs.map(s=>s.svg).join('')}</svg>
    <div style="flex:1;min-width:180px">
      ${segs.map(s => `<div style="display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid var(--border);font-size:12px">
        <span><span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:${s.c};margin-right:7px"></span><strong>${esc(s.name)}</strong> (${fmt(s.orders)})</span>
        <span>${fmt(s.net)} đ · <strong>${s.pct}%</strong></span>
      </div>`).join('')}
    </div>
  </div>`;
}

function renderOpsChart(comp, ret, cancel) {
  const tot = Math.max(1, comp + ret + cancel);
  document.getElementById('chartOps').innerHTML = [
    ['Hoàn tất (COMPLETED)', comp, '#3D7A52'],
    ['Hoàn trả (RETURNED)', ret, '#657DC4'],
    ['Đơn hủy (EXCLUDED)', cancel, '#C84B4B']
  ].map(([lbl, n, c]) => `
    <div style="margin-bottom:10px;font-size:12px">
      <div style="display:flex;justify-content:space-between"><span>${lbl}</span><strong>${fmt(n)} (${((n/tot)*100).toFixed(1)}%)</strong></div>
      <div style="height:9px;background:var(--soft);border-radius:99px;margin-top:3px"><div style="height:100%;width:${Math.max(2,Math.round(n/tot*100))}%;background:${c};border-radius:99px"></div></div>
    </div>
  `).join('');
}

function renderDiscountChart(entries) {
  const el = document.getElementById('chartDiscount');
  if (!DATA.schema.hasDiscount) { el.innerHTML = '<div class="missing-box">Bộ dữ liệu không ghi nhận chiết khấu (> 0 đ).</div>'; return; }
  el.innerHTML = `<table><thead><tr><th>Chi nhánh</th><th class="num">GMV</th><th class="num">Chiết khấu</th><th class="num">Tỷ lệ CK</th></tr></thead><tbody>` +
    entries.slice(0, 6).map(([name, v]) => `<tr><td>${esc(name)}</td><td class="num">${fmt(v.gmv)} đ</td><td class="num">${fmt(v.disc)} đ</td><td class="num">${v.gmv>0 ? ((v.disc/v.gmv)*100).toFixed(1) : '0.0'}%</td></tr>`).join('') +
    `</tbody></table>`;
}

function renderCustomerSection() {
  const el = document.getElementById('custBox');
  if (!DATA.schema.hasCustomer || !DATA.customers.length) {
    el.innerHTML = '<div class="missing-box">Chưa đủ dữ liệu để thực hiện phân tích Khách hàng (cần bổ sung trường Khách hàng / Mã KH).</div>';
    return;
  }
  el.innerHTML = `<div style="font-size:12.5px;margin-bottom:10px">Tổng khách hàng duy nhất: <strong>${fmt(DATA.meta.totalCustomers)}</strong> · Khách hàng quay lại (&gt;1 đơn): <strong>${fmt(DATA.meta.repeatCustomers)}</strong></div>
  <table><thead><tr><th>Top Khách hàng</th><th class="num">Số đơn</th><th class="num">Sản lượng</th><th class="num">Doanh thu thuần</th></tr></thead><tbody>` +
  DATA.customers.map(c => `<tr><td><strong>${esc(c.name)}</strong></td><td class="num">${fmt(c.orders)}</td><td class="num">${fmt(c.units)}</td><td class="num">${fmt(c.net)} đ</td></tr>`).join('') + `</tbody></table>`;
}

function renderAlerts() {
  const ins = (DATA.insights || []).map(r => {
    const cls = r.priority === 'Cao' ? 'critical' : r.priority === 'Trung bình' ? 'warning' : 'positive';
    return `<div class="alert ${cls}"><strong>[${esc(r.category)} · ${esc(r.priority)}] ${esc(r.title)}:</strong> ${esc(r.content)}</div>`;
  }).join('');
  const anom = (DATA.anomalies || []).map(a => `<div class="alert warning"><strong>${esc(a.title)}:</strong> ${esc(a.evidence)} — <em>${esc(a.cause)}</em></div>`).join('');
  document.getElementById('alertsBox').innerHTML = ins + anom || '<div class="missing-box">Chưa có cảnh báo.</div>';
}

function sortExp(col) { if (expSort === col) expDir *= -1; else { expSort = col; expDir = -1; } renderExplorer(); }
function renderExplorer() {
  const q = (document.getElementById('expSearch').value || '').toLowerCase();
  let list = q ? filteredRows.filter(r => r.d.includes(q) || r.s.toLowerCase().includes(q) || r.p.toLowerCase().includes(q) || r.c.toLowerCase().includes(q)) : [...filteredRows];
  list.sort((a,b) => (a[expSort] > b[expSort] ? 1 : -1) * expDir);
  const ps = 25, totalPages = Math.max(1, Math.ceil(list.length / ps));
  if (expPage > totalPages) expPage = totalPages;
  const pageItems = list.slice((expPage - 1) * ps, expPage * ps);
  document.getElementById('expInfo').textContent = `Hiển thị ${pageItems.length} / ${fmt(list.length)} nhóm dòng`;
  document.getElementById('expPageTxt').textContent = `Trang ${expPage} / ${totalPages}`;
  document.getElementById('expBody').innerHTML = pageItems.map(r => `
    <tr><td>${r.d}</td><td>${esc(r.s)}</td><td>${esc(r.p)}</td><td>${esc(r.c)}</td><td class="num">${fmt(r.q)}</td><td class="num">${fmt(r.net)} đ</td><td class="num">${DATA.schema.hasCost ? fmt(r.prof) + ' đ' : '—'}</td><td>${r.st}</td></tr>
  `).join('') || '<tr><td colspan="8" style="text-align:center;color:var(--brown-soft)">Không có dữ liệu.</td></tr>';
}

init();
</script>
</body>
</html>
""");
        return sb.ToString();
    }
}
