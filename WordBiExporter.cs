using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class WordBiExporter
{
    private const string FONT_NAME = "Times New Roman";
    private const string COLOR_PRIMARY = "3C4430";
    private const string COLOR_ACCENT = "56613F";
    private const string COLOR_MUTED = "555555";
    private const string COLOR_BORDER = "CCCCCC";
    private const string COLOR_TH_BG = "56613F";
    private const string COLOR_ALT_BG = "F4F4F1";

    public static byte[] BuildReport(DataAnalystHub hub, DeepAnalysisResult analysis)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            // Settings: UpdateFieldsOnOpen = true for automatic TOC
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(new UpdateFieldsOnOpen { Val = true });
            settingsPart.Settings.Save();

            // Styles
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = BuildStyles();
            stylesPart.Styles.Save();

            // Header & Footer
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(
                MakePara("Ngọc Linh Data Studio — Báo cáo Phân tích Dữ liệu Kinh doanh (BI)", 10, false, false, JustificationValues.Right, "888888")
            );
            headerPart.Header.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            var fPara = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Right }));
            fPara.Append(
                MakeRun("Báo cáo BI nội bộ  |  Trang ", 10, false, false, "888888"),
                new SimpleField { Instruction = "PAGE" },
                MakeRun(" / ", 10, false, false, "888888"),
                new SimpleField { Instruction = "NUMPAGES" }
            );
            footerPart.Footer = new Footer(fPara);
            footerPart.Footer.Save();

            // Helper metrics
            var m = analysis.OverallMetrics;
            var p = analysis.PricingAnalysis;
            var st = hub.LastStats;
            string dateRange = analysis.TimeSeries.Count > 0
                ? $"{analysis.TimeSeries.First().Label} – {analysis.TimeSeries.Last().Label}"
                : DateTime.Now.ToString("dd/MM/yyyy");
            string reportDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

            var sch = analysis.Schema;

            // ================= PHẦN 1. TRANG BÌA =================
            body.Append(MakePara("", 13));
            body.Append(MakePara("NGỌC LINH DATA STUDIO", 16, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara("EXECUTIVE BUSINESS INTELLIGENCE REPORT", 13, true, false, JustificationValues.Center, COLOR_ACCENT));
            body.Append(MakePara("BI Data Cleansing & Analytics Workbench", 11, false, true, JustificationValues.Center, COLOR_MUTED));
            body.Append(MakePara("", 22));
            body.Append(MakePara("BÁO CÁO PHÂN TÍCH KINH DOANH", 22, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara("VÀ HIỆU QUẢ HOẠT ĐỘNG", 22, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara("", 14));
            body.Append(MakePara($"Đơn vị / Bộ dữ liệu: {hub.SourceName}", 13, true, false, JustificationValues.Center));
            body.Append(MakePara($"Thời gian phân tích: {dateRange}", 13, false, false, JustificationValues.Center));
            body.Append(MakePara($"Quy mô dữ liệu: {m.TotalOrders:N0} bản ghi hợp lệ / {st.RawCount:N0} bản ghi đầu vào", 13, false, false, JustificationValues.Center));
            body.Append(MakePara($"Phiên bản báo cáo: #{hub.DataVersion} · Trạng thái: {(hub.IsHumanVerified ? "Đã kiểm duyệt" : "Chờ duyệt")}", 12, false, true, JustificationValues.Center, COLOR_ACCENT));
            body.Append(MakePara("", 30));
            body.Append(MakePara("Đơn vị thực hiện: Ngọc Linh Data Studio", 13, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara($"Ngày lập báo cáo: {reportDate}", 13, false, true, JustificationValues.Center, COLOR_MUTED));
            body.Append(PageBreak());

            // ================= MỤC LỤC TỰ ĐỘNG =================
            body.Append(MakePara("MỤC LỤC", 15, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(new Paragraph(new SimpleField { Instruction = "TOC \\o \"1-3\" \\h \\z \\u" }));
            body.Append(MakePara("(Nhấp chuột phải vào Mục lục và chọn 'Update Field' trong Microsoft Word để cập nhật số trang)", 10, false, true, JustificationValues.Center, "777777"));
            body.Append(PageBreak());

            // ================= PHẦN 2. EXECUTIVE SUMMARY =================
            body.Append(Heading1("PHẦN 2. EXECUTIVE SUMMARY – TÓM TẮT ĐIỀU HÀNH"));
            body.Append(BodyPara($"Báo cáo tổng hợp kết quả phân tích từ bộ dữ liệu '{hub.SourceName}' ({dateRange}), phiên bản #{hub.DataVersion}. Hệ thống đã chuẩn hóa {st.RawCount:N0} bản ghi đầu vào còn {m.TotalOrders:N0} giao dịch hợp lệ (loại trừ {st.DuplicatesRemoved:N0} đơn trùng lặp, {st.CancelledExcluded:N0} đơn hủy/lỗi, hạch toán giảm trừ {st.ReturnsAdjusted:N0} đơn hoàn trả), ngăn chặn {m.PreventedOvercountValue:N0} đ doanh thu ảo."));

            var profitExec = m.HasCostData
                ? $"{m.TotalGrossProfit:N0} đ (Biên LN {m.GrossMarginPercent}%)"
                : "Chưa có dữ liệu giá vốn thực tế";
            body.Append(CaptionPara("Bảng 1: Bảng tóm tắt chỉ số điều hành trọng yếu"));
            body.Append(BuildTable(
                new[] { "Chỉ tiêu điều hành", "Giá trị thực tế", "Ghi chú kiểm soát" },
                new[] {
                    new[] { "Tổng doanh thu gộp (GMV)", $"{m.TotalGrossRevenue:N0} đ", $"Chiết khấu: {m.TotalDiscountAmount:N0} đ ({p.AverageDiscountRate}%)" },
                    new[] { "Doanh thu thuần thực tế", $"{m.TotalNetRevenue:N0} đ", $"Trên {m.TotalOrders:N0} bản ghi sạch" },
                    new[] { "Lợi nhuận gộp & Biên LN", profitExec, m.HasCostData ? $"Tổng giá vốn: {m.TotalCost:N0} đ" : "Không tự ước tính giá vốn thiếu" },
                    new[] { "Giá trị đơn trung bình (AOV)", $"{p.AverageOrderValue:N0} đ", $"Tổng sản lượng: {analysis.TotalUnitsSold:N0} đơn vị" },
                    new[] { "Tỷ lệ hoàn trả / Hủy đơn", $"{m.ReturnRatePercent}% / {m.CancelRatePercent}%", $"{m.ReturnOrders:N0} đơn hoàn · {st.CancelledExcluded:N0} đơn hủy" }
                },
                new[] { 3000, 2600, 3400 }
            ));

            var topStore = analysis.StorePerformance.FirstOrDefault();
            var topProd = analysis.ProductPerformance.FirstOrDefault();
            var worstStore = analysis.StorePerformance.OrderByDescending(s => s.ReturnRatePercent).FirstOrDefault();

            body.Append(Heading2("Giải đáp 5 câu hỏi trọng tâm cho Ban điều hành"));
            body.Append(BulletPara($"1. Doanh nghiệp đang hoạt động như thế nào? Doanh thu thuần thực tế đạt {m.TotalNetRevenue:N0} đ trên {m.TotalOrders:N0} giao dịch hợp lệ; {(sch.HasStore && topStore != null ? $"chi nhánh dẫn đầu là '{topStore.StoreName}' ({topStore.NetRevenue:N0} đ)" : "hoạt động tập trung trên toàn bộ dữ liệu")}."));
            body.Append(BulletPara($"2. Kết quả kinh doanh đang tốt hay xấu? {(analysis.GrowthRatePercent >= 0 ? $"Xu hướng doanh thu ghi nhận mức tăng trưởng {analysis.GrowthRatePercent}% qua {analysis.TimeSeries.Count} ngày giao dịch" : $"Xu hướng doanh thu biến động giảm {analysis.GrowthRatePercent}% qua {analysis.TimeSeries.Count} ngày")}; tỷ lệ hoàn trả hiện ở mức {m.ReturnRatePercent}%, tỷ lệ hủy đơn {m.CancelRatePercent}%."));
            body.Append(BulletPara($"3. Vấn đề lớn nhất nằm ở đâu? {(worstStore != null && worstStore.ReturnRatePercent > 5 ? $"Chi nhánh '{worstStore.StoreName}' có tỷ lệ hoàn trả cao nhất ({worstStore.ReturnRatePercent}%, {worstStore.ReturnCount} đơn hoàn)" : st.DuplicatesRemoved + st.CancelledExcluded > 0 ? $"Dữ liệu thô phát sinh {st.DuplicatesRemoved + st.CancelledExcluded} bản ghi trùng lặp và đơn hủy gây sai lệch {m.PreventedOvercountValue:N0} đ" : "Cần tiếp tục theo dõi biên lợi nhuận và tỷ lệ chiết khấu bình quân ({p.AverageDiscountRate}%)")}."));
            body.Append(BulletPara($"4. Nguyên nhân nào có bằng chứng dữ liệu hỗ trợ? Toàn bộ kết luận dựa trên dữ liệu đã đối soát: loại {st.CancelledExcluded} đơn hủy, giảm trừ {st.ReturnsAdjusted} đơn hoàn và ghi nhận {analysis.Anomalies.Count} điểm biến động bất thường."));
            body.Append(BulletPara($"5. Ban điều hành nên ưu tiên hành động gì? (1) Kiểm soát điểm nóng hoàn trả/hủy đơn; (2) Tối ưu phân bổ hàng hóa tại {(topProd != null ? $"SKU dẫn đầu '{topProd.ProductName}'" : "các nhóm hàng chủ lực")}; (3) Chuẩn hóa dữ liệu giá vốn và chiết khấu trước mọi kỳ báo cáo."));

            // ================= PHẦN 3. TỔNG QUAN DỮ LIỆU VÀ PHƯƠNG PHÁP =================
            body.Append(Heading1("PHẦN 3. TỔNG QUAN DỮ LIỆU VÀ PHƯƠNG PHÁP PHÂN TÍCH"));
            body.Append(Heading2("3.1. Data Quality Assessment – Đánh giá chất lượng & độ tin cậy dữ liệu"));
            body.Append(BodyPara($"Nguồn dữ liệu: {hub.SourceName} ({hub.UploadedFiles.Count} tệp). Thời gian thu thập: {dateRange}. Mô hình phân tích: Tự động nhận diện cấu trúc trường dữ liệu (Data Schema Recognition)."));
            body.Append(CaptionPara("Bảng 2: Phân tách Dữ liệu gốc – Dữ liệu sau làm sạch – Dữ liệu phân tích"));
            body.Append(BuildTable(
                new[] { "Giai đoạn dữ liệu", "Số lượng / Giá trị", "Quy tắc xử lý & Mức độ tin cậy" },
                new[] {
                    new[] { "1. Dữ liệu gốc (Raw Data)", $"{st.RawCount:N0} dòng", "Tiếp nhận nguyên bản, không ghi đè dữ liệu gốc" },
                    new[] { "2. Bản ghi trùng lặp tách riêng", $"{st.DuplicatesRemoved:N0} dòng", "Khử trùng theo khóa (Mã đơn + SKU + Sản phẩm)" },
                    new[] { "3. Đơn hủy / lỗi thanh toán", $"{st.CancelledExcluded:N0} dòng", "Loại khỏi doanh thu thực tế (CANCELLED / FAILED)" },
                    new[] { "4. Dữ liệu thiếu / sai định dạng", $"{st.MissingOrInvalid:N0} dòng", "Tách bảng ngoại lệ; cho phép chỉnh sửa trực tiếp" },
                    new[] { "5. Đơn hoàn trả đối soát", $"{st.ReturnsAdjusted:N0} dòng", "Ghi nhận giảm trừ doanh thu thực tế" },
                    new[] { "6. Dòng thiếu giá vốn đơn vị", $"{st.ImputedCount:N0} dòng", "Để trống (0 đ)" },
                    new[] { "7. Dữ liệu sạch đưa vào phân tích", $"{st.CleanedCount:N0} dòng", hub.IsHumanVerified ? "Đã kiểm duyệt & phê duyệt chính thức" : "Đã làm sạch" }
                },
                new[] { 3200, 2200, 3600 }
            ));

            // ================= PHẦN 4. HỆ THỐNG KPI =================
            body.Append(Heading1("PHẦN 4. HỆ THỐNG KPI VÀ HIỆU QUẢ KINH DOANH"));
            body.Append(BodyPara("Hệ thống chỉ tính toán các chỉ số KPI khi dữ liệu đầu vào có đầy đủ trường thông tin tương ứng:"));
            var kpiRows = new List<string[]>
            {
                new[] { "Tổng doanh thu gộp (GMV)", $"{m.TotalGrossRevenue:N0}", "VNĐ", "∑(SL × Đơn giá) đơn bán", "Đo quy mô doanh số trước giảm giá" },
                new[] { "Doanh thu thuần thực tế", $"{m.TotalNetRevenue:N0}", "VNĐ", "GMV - Chiết khấu - Hoàn trả", "Doanh thu thực nhận cốt lõi" },
                new[] { "Tổng số đơn / Hoàn tất", $"{m.TotalOrders:N0} / {analysis.CompletedOrders:N0}", "Đơn", "Đếm bản ghi hợp lệ", "Quy mô giao dịch thực tế" },
                new[] { "Giá trị đơn TB (AOV)", $"{p.AverageOrderValue:N0}", "VNĐ", "DT thuần / Số đơn bán", "Đo giá trị giỏ hàng trung bình" },
                new[] { "Tỷ lệ hoàn trả / Hủy đơn", $"{m.ReturnRatePercent}% / {m.CancelRatePercent}%", "%", "Đơn hoàn hoặc hủy / Tổng đơn", "Đánh giá chất lượng vận hành" }
            };
            if (m.HasCostData)
            {
                kpiRows.Add(new[] { "Tổng giá vốn (COGS)", $"{m.TotalCost:N0}", "VNĐ", "∑(SL × Giá vốn ĐV)", "Chi phí hàng bán trực tiếp" });
                kpiRows.Add(new[] { "Lợi nhuận gộp", $"{m.TotalGrossProfit:N0}", "VNĐ", "DT thuần - Giá vốn", "Hiệu quả sinh lời cốt lõi" });
                kpiRows.Add(new[] { "Biên lợi nhuận gộp", $"{m.GrossMarginPercent}%", "%", "(LN gộp / DT thuần) × 100%", "Tỷ suất sinh lời trên doanh thu" });
            }
            if (sch.HasCustomer)
                kpiRows.Add(new[] { "Tổng khách hàng / Quay lại", $"{analysis.TotalCustomers:N0} / {analysis.RepeatCustomers:N0}", "KH", "Đếm KH duy nhất & >1 đơn", "Đánh giá độ trung thành khách hàng" });

            body.Append(CaptionPara("Bảng 3: Hệ thống chỉ số KPI được kích hoạt theo cấu trúc dữ liệu"));
            body.Append(BuildTable(
                new[] { "Tên chỉ số KPI", "Giá trị", "Đơn vị", "Công thức tính", "Ý nghĩa quản trị" },
                kpiRows.ToArray(),
                new[] { 2000, 1400, 800, 2300, 2500 }
            ));

            // ================= PHẦN 5. PHÂN TÍCH CHUYÊN SÂU =================
            body.Append(Heading1("PHẦN 5. PHÂN TÍCH CHUYÊN SÂU HOẠT ĐỘNG KINH DOANH"));

            body.Append(Heading2("5.1. Phân tích Doanh thu, Lợi nhuận và Xu hướng thời gian"));
            body.Append(BodyPara($"Tổng doanh thu gộp đạt {m.TotalGrossRevenue:N0} đ; doanh thu thuần thực tế đạt {m.TotalNetRevenue:N0} đ. {(m.HasCostData ? $"Tổng giá vốn {m.TotalCost:N0} đ, lợi nhuận gộp đạt {m.TotalGrossProfit:N0} đ (biên {m.GrossMarginPercent}%)." : "Chưa có dữ liệu giá vốn thực tế nên hệ thống không tính lợi nhuận gộp ước tính.")}"));
            if (analysis.TimeSeries.Count >= 2)
            {
                body.Append(BodyPara($"Chuỗi thời gian ghi nhận qua {analysis.TimeSeries.Count} ngày ({dateRange}), tốc độ tăng trưởng kỳ đạt {analysis.GrowthRatePercent}%."));
                var chart1Bytes = RenderLineChartPng("Xu hướng Doanh thu thuần theo thời gian (VNĐ)", analysis.TimeSeries);
                body.Append(ImagePara(mainPart, chart1Bytes, "chart_trend.png", 5400000L, 2500000L));
                body.Append(CaptionPara("Hình 1: Biểu đồ xu hướng Doanh thu thuần theo thời gian (VNĐ)"));
            }

            body.Append(Heading2("5.2. Phân tích Sản phẩm"));
            if (sch.HasProduct && analysis.ProductPerformance.Count > 0)
            {
                body.Append(CaptionPara("Bảng 4: Top sản phẩm theo Doanh thu thuần"));
                body.Append(BuildTable(
                    new[] { "SKU", "Tên sản phẩm", "SL bán", "Doanh thu thuần (đ)", "Lợi nhuận gộp (đ)", "Biên LN %" },
                    analysis.ProductPerformance.Take(10).Select(pr => new[] {
                        pr.Sku, pr.ProductName, $"{pr.QuantitySold:N0}", $"{pr.NetRevenue:N0}", m.HasCostData ? $"{pr.GrossProfit:N0}" : "N/A", m.HasCostData ? $"{pr.MarginPercent}%" : "N/A"
                    }).ToArray(),
                    new[] { 1200, 2800, 900, 1700, 1500, 900 }
                ));
                var chart3Bytes = RenderBarChartPng("Top Sản phẩm theo Doanh thu thuần (VNĐ)",
                    analysis.ProductPerformance.Take(6).Select(pr => (pr.ProductName, (double)pr.NetRevenue)).ToList(),
                    System.Drawing.Color.FromArgb(102, 79, 65));
                body.Append(ImagePara(mainPart, chart3Bytes, "chart_products.png", 5400000L, 2500000L));
                body.Append(CaptionPara("Hình 2: Top sản phẩm đóng góp doanh thu cao nhất (VNĐ)"));
            }
            else
            {
                body.Append(BodyPara("Chưa đủ dữ liệu để thực hiện phân tích chi tiết sản phẩm (cần bổ sung trường Tên sản phẩm / SKU)."));
            }

            body.Append(Heading2("5.3. Phân tích Chi nhánh và Khu vực"));
            if (sch.HasStore && analysis.StorePerformance.Count > 0)
            {
                body.Append(CaptionPara("Bảng 5: Hiệu quả kinh doanh theo Chi nhánh"));
                body.Append(BuildTable(
                    new[] { "Chi nhánh", "Số đơn", "Doanh thu thuần (đ)", "Lợi nhuận gộp (đ)", "Biên LN %", "Tỷ lệ hoàn %", "Đánh giá" },
                    analysis.StorePerformance.Select(s => new[] {
                        s.StoreName, $"{s.OrderCount:N0}", $"{s.NetRevenue:N0}", m.HasCostData ? $"{s.GrossProfit:N0}" : "N/A", m.HasCostData ? $"{s.MarginPercent}%" : "N/A", $"{s.ReturnRatePercent}%", s.StatusAssessment
                    }).ToArray(),
                    new[] { 2000, 800, 1600, 1500, 1000, 1000, 1100 }
                ));
                var chart2Bytes = RenderBarChartPng("So sánh Doanh thu thuần theo Chi nhánh (VNĐ)",
                    analysis.StorePerformance.Take(8).Select(s => (s.StoreName, (double)s.NetRevenue)).ToList(),
                    System.Drawing.Color.FromArgb(221, 158, 94));
                body.Append(ImagePara(mainPart, chart2Bytes, "chart_stores.png", 5400000L, 2500000L));
                body.Append(CaptionPara("Hình 3: So sánh Doanh thu thuần giữa các chi nhánh (VNĐ)"));
            }
            else
            {
                body.Append(BodyPara("Chưa đủ dữ liệu để thực hiện phân tích chi nhánh (cần bổ sung trường Chi nhánh / Cửa hàng)."));
            }

            body.Append(Heading2("5.4. Phân tích Kênh bán hàng, Khách hàng & Chiết khấu"));
            if (sch.HasChannel && analysis.ChannelBreakdown.Count > 0)
            {
                body.Append(CaptionPara("Bảng 6: Cơ cấu doanh thu theo Kênh bán hàng"));
                body.Append(BuildTable(
                    new[] { "Kênh bán", "Số đơn", "Doanh thu thuần (đ)", "Tỷ trọng %" },
                    analysis.ChannelBreakdown.Select(c => new[] { c.ChannelName, $"{c.OrderCount:N0}", $"{c.NetRevenue:N0}", $"{c.RevenueSharePercent}%" }).ToArray(),
                    new[] { 2500, 1500, 3000, 2000 }
                ));
            }
            else
            {
                body.Append(BodyPara("Chưa đủ dữ liệu để phân tách Kênh bán hàng (cần bổ sung trường Kênh / Channel)."));
            }

            if (sch.HasCustomer && analysis.CustomerPerformance.Count > 0)
            {
                body.Append(BodyPara($"Phân tích Khách hàng: Ghi nhận {analysis.TotalCustomers:N0} khách hàng duy nhất, trong đó {analysis.RepeatCustomers:N0} khách hàng mua lặp lại (> 1 giao dịch)."));
                body.Append(CaptionPara("Bảng 7: Top khách hàng đóng góp doanh thu cao nhất"));
                body.Append(BuildTable(
                    new[] { "Khách hàng", "Số đơn", "Sản lượng", "Doanh thu thuần (đ)" },
                    analysis.CustomerPerformance.Take(8).Select(c => new[] { c.CustomerName, $"{c.OrderCount:N0}", $"{c.TotalUnits:N0}", $"{c.NetRevenue:N0}" }).ToArray(),
                    new[] { 3400, 1400, 1400, 2800 }
                ));
            }
            else
            {
                body.Append(BodyPara("Chưa đủ dữ liệu để thực hiện phân tích Khách hàng (cần bổ sung trường Mã / Tên khách hàng)."));
            }

            if (sch.HasDiscount)
                body.Append(BodyPara($"Chính sách Chiết khấu: Có {p.DiscountedOrderCount:N0} đơn áp dụng giảm giá, tổng chiết khấu {m.TotalDiscountAmount:N0} đ (tỷ lệ bình quân {p.AverageDiscountRate}%, cao nhất {p.MaxDiscountPercent}%)."));

            // ================= PHẦN 6. BUSINESS INSIGHTS =================
            body.Append(Heading1("PHẦN 6. BUSINESS INSIGHTS – PHÁT HIỆN VẤN ĐỀ VÀ CƠ HỘI"));
            int insIdx = 1;
            foreach (var rec in analysis.StructuredRecommendations)
            {
                string sevTag = rec.Priority == "Cao" ? "Critical / High" : rec.Priority == "Trung bình" ? "Medium / Opportunity" : "Medium";
                body.Append(Heading2($"Insight {insIdx++} [{sevTag}]: {rec.Title}"));
                body.Append(BulletPara($"1. Phát hiện: Ghi nhận tại nhóm '{rec.Category}' (cập nhật ngày {rec.DetectedDate})."));
                body.Append(BulletPara($"2. Bằng chứng: {rec.Content}"));
                body.Append(BulletPara("3. Phân tích: Chênh lệch xuất hiện trực tiếp trên dữ liệu sạch sau khi đã loại bỏ đơn trùng và đơn hủy."));
                body.Append(BulletPara("4. Tác động: Ảnh hưởng trực tiếp đến doanh thu thuần thực tế, hiệu quả chi nhánh và chất lượng dòng tiền."));
                body.Append(BulletPara("5. Nguyên nhân: Được chứng minh từ số liệu giao dịch thực tế; cần đối chiếu thêm định mức vận hành tại điểm bán để xác minh nguyên nhân vận hành gốc."));
                body.Append(BulletPara($"6. Hành động: Rà soát và áp dụng biện pháp kiểm soát chuyên biệt cho nhóm {rec.Category}."));
            }

            // ================= PHẦN 7. KHUYẾN NGHỊ CHIẾN LƯỢC =================
            body.Append(Heading1("PHẦN 7. KHUYẾN NGHỊ CHIẾN LƯỢC VÀ GIẢI PHÁP QUẢN TRỊ"));
            var highPriorityStores = analysis.StorePerformance.Where(s => (m.HasCostData && s.MarginPercent < 12) || s.ReturnRatePercent > 10).ToList();
            var solutions = new[]
            {
                new {
                    Level = "Giải pháp 1 (Ưu tiên Critical/High)",
                    Title = "Kiểm soát dữ liệu ngoại lệ, đơn hủy & trần chiết khấu",
                    Problem = $"Phát hiện {st.DuplicatesRemoved + st.CancelledExcluded} bản ghi trùng/hủy (ngăn chặn {m.PreventedOvercountValue:N0} đ doanh thu ảo) và tỷ lệ CK bình quân {p.AverageDiscountRate}%.",
                    Action = "Thiết lập cổng kiểm duyệt dữ liệu 4 bước trên Ngọc Linh Data Studio và quy chế phê duyệt khi chiết khấu vượt ngưỡng.",
                    Owner = "Giám đốc Tài chính (CFO) & Bộ phận BI",
                    Kpi = "Sai lệch dữ liệu = 0%; Duy trì tỷ lệ chiết khấu hợp lý",
                    Timeline = "0–30 ngày",
                    Resources = "Quy trình đối soát BI & Quy chế giá"
                },
                new {
                    Level = "Giải pháp 2 (Ưu tiên High)",
                    Title = highPriorityStores.Count > 0 ? $"Xử lý điểm nóng hoàn trả/biên thấp tại {string.Join(", ", highPriorityStores.Select(s => s.StoreName))}" : "Tối ưu cơ cấu sản phẩm & chi nhánh chủ lực",
                    Problem = $"Tỷ lệ hoàn trả toàn hệ thống đạt {m.ReturnRatePercent}% ({m.ReturnOrders} đơn hoàn).",
                    Action = "Kiểm tra quy trình đóng gói, xác nhận đơn trước giao và tái phân bổ hàng hóa cho nhóm SKU dẫn đầu.",
                    Owner = "Giám đốc Vận hành (COO) & Quản lý Chi nhánh",
                    Kpi = "Tỷ lệ hoàn trả (%); Doanh thu thuần theo chi nhánh",
                    Timeline = "31–90 ngày",
                    Resources = "Dữ liệu đơn hoàn & Quy trình QC"
                },
                new {
                    Level = "Giải pháp 3 (Ưu tiên Medium)",
                    Title = "Chuẩn hóa từ điển dữ liệu (bao gồm giá vốn thực tế)",
                    Problem = st.ImputedCount > 0 ? $"Còn {st.ImputedCount:N0} dòng thiếu thông tin giá vốn đơn vị thực tế." : "Duy trì chuẩn dữ liệu đồng bộ giữa các chi nhánh và kênh bán.",
                    Action = "Chuẩn hóa mẫu nhập liệu đồng nhất có đủ Mã đơn, Ngày, Chi nhánh, Sản phẩm, Giá vốn thực.",
                    Owner = "Phòng Kế toán & Bộ phận Dữ liệu",
                    Kpi = "Độ phủ trường dữ liệu chuẩn = 100%",
                    Timeline = "Trên 90 ngày",
                    Resources = "Từ điển dữ liệu chuẩn BI"
                }
            };
            foreach (var sol in solutions)
            {
                body.Append(Heading2($"{sol.Level}: {sol.Title}"));
                body.Append(BulletPara($"Vấn đề & Căn cứ số liệu: {sol.Problem}"));
                body.Append(BulletPara($"Hành động cụ thể: {sol.Action}"));
                body.Append(BulletPara($"Bộ phận chịu trách nhiệm: {sol.Owner} | Thời gian: {sol.Timeline}"));
                body.Append(BulletPara($"KPI đánh giá: {sol.Kpi} | Nguồn lực: {sol.Resources}"));
            }

            // ================= PHẦN 8. KẾ HOẠCH HÀNH ĐỘNG =================
            body.Append(Heading1("PHẦN 8. KẾ HOẠCH HÀNH ĐỘNG"));
            body.Append(CaptionPara("Bảng 8: Lộ trình hành động 3 giai đoạn cho Ban điều hành"));
            body.Append(BuildTable(
                new[] { "Giai đoạn", "Mục tiêu & Hành động trọng tâm", "Bộ phận phụ trách", "KPI theo dõi" },
                new[] {
                    new[] { "Ngắn hạn (0–30 ngày)", "Khóa cổng kiểm duyệt dữ liệu; xử lý triệt để đơn trùng/hủy và kiểm soát chiết khấu vượt ngưỡng.", "CFO, Phòng Kinh doanh & BI", "Sai lệch = 0%; CK TB %" },
                    new[] { "Trung hạn (31–90 ngày)", "Rà soát chi nhánh có tỷ lệ hoàn cao; tối ưu tồn kho & phân bổ SKU chủ lực.", "COO & Quản lý Chi nhánh", "Tỷ lệ hoàn %; Doanh thu chi nhánh" },
                    new[] { "Dài hạn (> 90 ngày)", "Đồng bộ 100% giá vốn thực tế và tự động hóa báo cáo điều hành định kỳ.", "Kế toán & Ban Giám đốc", "Độ phủ dữ liệu chuẩn = 100%" }
                },
                new[] { 1800, 3600, 1800, 1800 }
            ));

            // ================= PHẦN 9. KẾT LUẬN =================
            body.Append(Heading1("PHẦN 9. KẾT LUẬN VÀ ĐỀ XUẤT CHO BAN ĐIỀU HÀNH"));
            body.Append(BodyPara($"Tổng hợp: Bộ dữ liệu '{hub.SourceName}' ghi nhận doanh thu thuần thực tế {m.TotalNetRevenue:N0} đ (GMV {m.TotalGrossRevenue:N0} đ) sau khi loại {st.DuplicatesRemoved + st.CancelledExcluded:N0} bản ghi sai lệch."));
            body.Append(Heading2("3 Quyết định quản trị cần ưu tiên"));
            body.Append(BulletPara("1. Áp dụng bắt buộc quy trình kiểm duyệt dữ liệu sạch trước khi chốt doanh thu kỳ báo cáo."));
            body.Append(BulletPara("2. Tập trung nguồn lực kiểm soát tỷ lệ hoàn trả và tối ưu hiệu suất tại các chi nhánh trọng điểm."));
            body.Append(BulletPara("3. Bổ sung đầy đủ trường giá vốn thực tế (nếu còn thiếu) để đo lường chính xác biên lợi nhuận gộp."));

            // ================= PHẦN 10. PHỤ LỤC =================
            body.Append(Heading1("PHẦN 10. PHỤ LỤC"));
            body.Append(BodyPara($"Nguồn tệp: {hub.SourceName} | Phiên bản #{hub.DataVersion} | Thời điểm tạo: {reportDate} | Số lần chỉnh sửa trực tiếp: {hub.EditLogs.Count}."));
            body.Append(CaptionPara("Bảng 9: Từ điển dữ liệu chuẩn hóa"));
            body.Append(BuildTable(
                new[] { "Tên trường", "Kiểu dữ liệu", "Vai trò trong phân tích BI" },
                new[] {
                    new[] { "OrderCode", "Text", "Khóa chính khử trùng lặp và đếm số lượng đơn hàng" },
                    new[] { "OrderDate", "Date", "Phân tích chuỗi thời gian và tốc độ tăng trưởng" },
                    new[] { "StoreName", "Text", "Phân tích so sánh hiệu quả giữa các chi nhánh" },
                    new[] { "ProductName / Sku", "Text", "Đánh giá top sản phẩm doanh thu và lợi nhuận" },
                    new[] { "NetRevenue / GrossProfit", "Currency", "KPI cốt lõi đo lường quy mô và khả năng sinh lời" },
                    new[] { "Status", "Text", "Loại trừ đơn hủy và hạch toán giảm trừ hàng hoàn" }
                },
                new[] { 2400, 1600, 5000 }
            ));

            // Section properties: A4 + margins + header/footer + titlePg
            var sectPr = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) },
                new PageSize { Width = 11906, Height = 16838 }, // A4
                new PageMargin { Top = 1440, Bottom = 1440, Right = 1134, Left = 1701, Header = 720, Footer = 720 },
                new TitlePage()
            );
            body.Append(sectPr);
            mainPart.Document.Save();
        }
        return ms.ToArray();
    }

    // ---------------- Chart PNG Renderers (System.Drawing.Common) ----------------
    private static byte[] RenderLineChartPng(string title, List<TimeSeriesPoint> points)
    {
        int w = 900, h = 400;
        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(System.Drawing.Color.White);

        using var titleFont = new System.Drawing.Font("Times New Roman", 13, FontStyle.Bold);
        using var labelFont = new System.Drawing.Font("Times New Roman", 9);
        using var brushDark = new SolidBrush(System.Drawing.Color.FromArgb(20, 20, 20));
        using var brushOlive = new SolidBrush(System.Drawing.Color.FromArgb(86, 97, 63));
        using var penGrid = new Pen(System.Drawing.Color.FromArgb(230, 230, 225), 1);
        using var penLine = new Pen(System.Drawing.Color.FromArgb(86, 97, 63), 3);

        g.DrawString(title, titleFont, brushDark, 24, 14);
        int pl = 95, pr = 30, pt = 55, pb = 55;
        int cw = w - pl - pr, ch = h - pt - pb;

        double maxVal = Math.Max(1, points.Max(p => (double)p.NetRevenue));
        double minVal = Math.Min(0, points.Min(p => (double)p.NetRevenue));
        double span = Math.Max(1, maxVal - minVal);

        for (int i = 0; i <= 4; i++)
        {
            float y = pt + ch * (1f - i / 4f);
            g.DrawLine(penGrid, pl, y, pl + cw, y);
            double v = minVal + span * i / 4.0;
            g.DrawString(FormatShort(v), labelFont, Brushes.Gray, 12, y - 8);
        }

        var pts = new PointF[points.Count];
        for (int i = 0; i < points.Count; i++)
        {
            float x = points.Count == 1 ? pl + cw / 2f : pl + (float)i * cw / (points.Count - 1);
            float y = pt + ch * (1f - (float)(((double)points[i].NetRevenue - minVal) / span));
            pts[i] = new PointF(x, y);
        }

        if (pts.Length > 1) g.DrawLines(penLine, pts);
        for (int i = 0; i < pts.Length; i++)
        {
            g.FillEllipse(brushOlive, pts[i].X - 4, pts[i].Y - 4, 8, 8);
            if (points.Count <= 12 || i % Math.Max(1, points.Count / 8) == 0)
                g.DrawString(points[i].Label, labelFont, brushDark, pts[i].X - 15, pt + ch + 10);
        }

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    private static byte[] RenderBarChartPng(string title, List<(string Label, double Value)> items, System.Drawing.Color barColor)
    {
        int w = 900, h = 400;
        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(System.Drawing.Color.White);

        using var titleFont = new System.Drawing.Font("Times New Roman", 13, FontStyle.Bold);
        using var labelFont = new System.Drawing.Font("Times New Roman", 9);
        using var brushDark = new SolidBrush(System.Drawing.Color.FromArgb(20, 20, 20));
        using var brushBar = new SolidBrush(barColor);
        using var penGrid = new Pen(System.Drawing.Color.FromArgb(230, 230, 225), 1);

        g.DrawString(title, titleFont, brushDark, 24, 14);
        int pl = 95, pr = 30, pt = 55, pb = 65;
        int cw = w - pl - pr, ch = h - pt - pb;

        double maxVal = Math.Max(1, items.Max(x => x.Value));
        for (int i = 0; i <= 4; i++)
        {
            float y = pt + ch * (1f - i / 4f);
            g.DrawLine(penGrid, pl, y, pl + cw, y);
            g.DrawString(FormatShort(maxVal * i / 4.0), labelFont, Brushes.Gray, 12, y - 8);
        }

        float slot = (float)cw / Math.Max(1, items.Count);
        float bw = Math.Min(65, slot * 0.55f);
        for (int i = 0; i < items.Count; i++)
        {
            float bh = (float)(Math.Max(0, items[i].Value) / maxVal * ch);
            float x = pl + i * slot + (slot - bw) / 2f;
            float y = pt + ch - bh;
            g.FillRectangle(brushBar, x, y, bw, bh);
            g.DrawString(FormatShort(items[i].Value), labelFont, brushDark, x, Math.Max(pt - 16, y - 16));
            string lbl = items[i].Label.Length > 15 ? items[i].Label[..14] + "…" : items[i].Label;
            g.DrawString(lbl, labelFont, brushDark, x - 4, pt + ch + 10);
        }

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    private static string FormatShort(double v)
    {
        double a = Math.Abs(v);
        if (a >= 1e9) return (v / 1e9).ToString("0.0") + " tỷ";
        if (a >= 1e6) return (v / 1e6).ToString("0.0") + " tr";
        return Math.Round(v).ToString("N0");
    }

    // ---------------- OpenXML Helpers ----------------
    private static Styles BuildStyles()
    {
        var styles = new Styles();
        styles.Append(new DocDefaults(
            new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { Ascii = FONT_NAME, HighAnsi = FONT_NAME, ComplexScript = FONT_NAME },
                new FontSize { Val = "26" }, // 13pt
                new FontSizeComplexScript { Val = "26" }
            )),
            new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                new SpacingBetweenLines { After = "120", Line = "324", LineRule = LineSpacingRuleValues.Auto } // 1.35 line spacing
            ))
        ));

        styles.Append(MakeHeadingStyle("Heading1", "heading 1", "32", COLOR_PRIMARY, "240", "120"));
        styles.Append(MakeHeadingStyle("Heading2", "heading 2", "28", COLOR_ACCENT, "180", "80"));
        styles.Append(MakeHeadingStyle("Heading3", "heading 3", "26", "141414", "120", "60"));
        return styles;
    }

    private static Style MakeHeadingStyle(string id, string name, string sz, string color, string before, string after) =>
        new(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new UIPriority { Val = 9 },
            new PrimaryStyle(),
            new StyleParagraphProperties(new KeepNext(), new SpacingBetweenLines { Before = before, After = after, Line = "324", LineRule = LineSpacingRuleValues.Auto }),
            new StyleRunProperties(new RunFonts { Ascii = FONT_NAME, HighAnsi = FONT_NAME }, new Bold(), new W.Color { Val = color }, new FontSize { Val = sz }, new FontSizeComplexScript { Val = sz })
        ) { Type = StyleValues.Paragraph, StyleId = id };

    private static Paragraph Heading1(string text) => new(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), MakeRun(text, 16, true, false, COLOR_PRIMARY));
    private static Paragraph Heading2(string text) => new(new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }), MakeRun(text, 14, true, false, COLOR_ACCENT));
    private static Paragraph BodyPara(string text) => MakePara(text, 13, false, false, JustificationValues.Both);
    private static Paragraph CaptionPara(string text) => MakePara(text, 12, true, true, JustificationValues.Center, COLOR_MUTED);
    private static Paragraph BulletPara(string text) => MakePara("•  " + text, 13, false, false, JustificationValues.Both);
    private static Paragraph PageBreak() => new(new Run(new Break { Type = BreakValues.Page }));

    private static Paragraph MakePara(string text, int pt, bool bold = false, bool italic = false, JustificationValues? align = null, string? hexColor = null)
    {
        var pp = new ParagraphProperties(new SpacingBetweenLines { After = "120", Line = "324", LineRule = LineSpacingRuleValues.Auto });
        if (align.HasValue) pp.Append(new Justification { Val = align.Value });
        return new Paragraph(pp, MakeRun(text, pt, bold, italic, hexColor));
    }

    private static Run MakeRun(string text, int pt, bool bold = false, bool italic = false, string? hexColor = null)
    {
        var rp = new RunProperties(
            new RunFonts { Ascii = FONT_NAME, HighAnsi = FONT_NAME, ComplexScript = FONT_NAME }
        );
        if (bold) rp.Append(new Bold());
        if (italic) rp.Append(new Italic());
        if (!string.IsNullOrEmpty(hexColor)) rp.Append(new W.Color { Val = hexColor });
        rp.Append(
            new FontSize { Val = (pt * 2).ToString() },
            new FontSizeComplexScript { Val = (pt * 2).ToString() }
        );
        return new Run(rp, new W.Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
    }

    private static W.Table BuildTable(string[] headers, string[][] rows, int[] colWidths)
    {
        var tbl = new W.Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER }
            )
        ));

        var grid = new TableGrid();
        foreach (var w in colWidths) grid.Append(new GridColumn { Width = w.ToString() });
        tbl.Append(grid);

        // Header row
        var thRow = new TableRow(new TableRowProperties(new TableHeader(), new CantSplit()));
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = colWidths[i].ToString(), Type = TableWidthUnitValues.Dxa },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = COLOR_TH_BG },
                    new TableCellMargin(new TopMargin { Width = "80", Type = TableWidthUnitValues.Dxa }, new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa }, new BottomMargin { Width = "80", Type = TableWidthUnitValues.Dxa }, new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa })
                ),
                MakePara(headers[i], 11, true, false, JustificationValues.Left, "FFFFFF")
            );
            thRow.Append(cell);
        }
        tbl.Append(thRow);

        for (int r = 0; r < rows.Length; r++)
        {
            var tr = new TableRow(new TableRowProperties(new CantSplit()));
            string bg = r % 2 == 1 ? COLOR_ALT_BG : "FFFFFF";
            for (int c = 0; c < headers.Length; c++)
            {
                string val = c < rows[r].Length ? rows[r][c] : "";
                var cell = new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Width = colWidths[c].ToString(), Type = TableWidthUnitValues.Dxa },
                        new Shading { Val = ShadingPatternValues.Clear, Fill = bg },
                        new TableCellMargin(new TopMargin { Width = "70", Type = TableWidthUnitValues.Dxa }, new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa }, new BottomMargin { Width = "70", Type = TableWidthUnitValues.Dxa }, new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa })
                    ),
                    MakePara(val, 11, false, false, JustificationValues.Left)
                );
                tr.Append(cell);
            }
            tbl.Append(tr);
        }
        return tbl;
    }

    private static Paragraph ImagePara(MainDocumentPart mainPart, byte[] pngBytes, string name, long cx, long cy)
    {
        var imgPart = mainPart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(pngBytes)) imgPart.FeedData(ms);
        string relId = mainPart.GetIdOfPart(imgPart);

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = (uint)Random.Shared.Next(100, 99999), Name = name },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
        );

        return new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }), new Run(drawing));
    }
}
