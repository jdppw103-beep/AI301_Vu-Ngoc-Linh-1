using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

public static class ExcelBiExporter
{
    // Style indices
    private const uint ST_DEFAULT = 0;
    private const uint ST_TITLE = 1;
    private const uint ST_SUBTITLE = 2;
    private const uint ST_TH = 3;
    private const uint ST_TEXT = 4;
    private const uint ST_TEXT_BOLD = 5;
    private const uint ST_NUM = 6;
    private const uint ST_CURRENCY = 7;
    private const uint ST_PCT = 8;
    private const uint ST_DATE = 9;
    private const uint ST_KPI_LABEL = 10;
    private const uint ST_KPI_VAL_CUR = 11;
    private const uint ST_KPI_VAL_NUM = 12;
    private const uint ST_KPI_VAL_PCT = 13;
    private const uint ST_NEG_CUR = 14;
    private const uint ST_SECTION = 15;

    public static byte[] BuildWorkbook(DataAnalystHub hub, DeepAnalysisResult analysis)
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = BuildStylesheet();
            stylesPart.Stylesheet.Save();

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            uint sheetId = 1;

            // 1. RAW DATA
            BuildRawDataSheet(wbPart, sheets, ref sheetId, hub);
            // 2. CLEANED DATA (Excel Table)
            int cleanRows = BuildCleanedDataSheet(wbPart, sheets, ref sheetId, hub);
            // 3. DATA DICTIONARY
            BuildDictionarySheet(wbPart, sheets, ref sheetId);
            // 4. CLEANING LOG
            BuildCleaningLogSheet(wbPart, sheets, ref sheetId, hub);
            // 5. ANALYSIS (Formulas referencing CLEANED DATA)
            var ranges = BuildAnalysisSheet(wbPart, sheets, ref sheetId, hub, analysis, cleanRows);
            // 6. DASHBOARD (KPI Cards + Charts + Insights)
            BuildDashboardSheet(wbPart, sheets, ref sheetId, hub, analysis, ranges);

            wbPart.Workbook.CalculationProperties = new CalculationProperties
            {
                CalculationId = 191029,
                FullCalculationOnLoad = true
            };
            wbPart.Workbook.Save();
        }
        return ms.ToArray();
    }

    private static void WriteSaxRowStart(OpenXmlWriter writer, uint rowIndex, double? height = null)
    {
        var attrs = new List<OpenXmlAttribute> { new("r", null, rowIndex.ToString()) };
        if (height.HasValue)
        {
            attrs.Add(new("ht", null, height.Value.ToString(CultureInfo.InvariantCulture)));
            attrs.Add(new("customHeight", null, "1"));
        }
        writer.WriteStartElement(new Row(), attrs);
    }

    private static void WriteSaxTextCell(OpenXmlWriter writer, string cellRef, string text, uint styleIndex)
    {
        writer.WriteStartElement(new Cell(), new[]
        {
            new OpenXmlAttribute("r", null, cellRef),
            new OpenXmlAttribute("t", null, "inlineStr"),
            new OpenXmlAttribute("s", null, styleIndex.ToString())
        });
        writer.WriteElement(new InlineString(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }));
        writer.WriteEndElement();
    }

    private static void WriteSaxNumCell(OpenXmlWriter writer, string cellRef, decimal val, uint styleIndex)
    {
        writer.WriteStartElement(new Cell(), new[]
        {
            new OpenXmlAttribute("r", null, cellRef),
            new OpenXmlAttribute("s", null, styleIndex.ToString())
        });
        writer.WriteElement(new CellValue(val.ToString(CultureInfo.InvariantCulture)));
        writer.WriteEndElement();
    }

    private static void WriteSaxFormulaNumCell(OpenXmlWriter writer, string cellRef, string formula, decimal cachedVal, uint styleIndex)
    {
        writer.WriteStartElement(new Cell(), new[]
        {
            new OpenXmlAttribute("r", null, cellRef),
            new OpenXmlAttribute("s", null, styleIndex.ToString())
        });
        writer.WriteElement(new CellFormula(formula));
        writer.WriteElement(new CellValue(cachedVal.ToString(CultureInfo.InvariantCulture)));
        writer.WriteEndElement();
    }

    private static void WriteSaxColumns(OpenXmlWriter writer, params double[] widths)
    {
        writer.WriteStartElement(new Columns());
        for (uint i = 0; i < widths.Length; i++)
        {
            writer.WriteElement(new Column { Min = i + 1, Max = i + 1, Width = widths[i], CustomWidth = true });
        }
        writer.WriteEndElement();
    }

    private static void BuildRawDataSheet(WorkbookPart wbPart, Sheets sheets, ref uint sheetId, DataAnalystHub hub)
    {
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        string[] headers = { "Mã đơn", "Ngày đơn", "Khách hàng", "Chi nhánh", "SKU", "Tên sản phẩm", "Số lượng", "Đơn giá", "Chiết khấu", "Giá vốn ĐV", "Trạng thái", "Kênh" };

        using (var writer = OpenXmlWriter.Create(wsPart))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement(new Worksheet());
            WriteSaxColumns(writer, 14, 14, 22, 24, 14, 28, 10, 16, 16, 16, 14, 14);

            writer.WriteStartElement(new SheetData());
            WriteSaxRowStart(writer, 1);
            WriteSaxTextCell(writer, "A1", "RAW DATA — DỮ LIỆU GỐC TỪ NGUỒN KHÁCH HÀNG", ST_TITLE);
            writer.WriteEndElement();

            WriteSaxRowStart(writer, 2);
            WriteSaxTextCell(writer, "A2", $"Nguồn: {hub.SourceName} | Số bản ghi gốc: {hub.RawData.Count} | Số trường: 12 | Thời điểm xuất: {DateTime.Now:dd/MM/yyyy HH:mm}", ST_SUBTITLE);
            writer.WriteEndElement();

            WriteSaxRowStart(writer, 4, 24);
            for (int c = 0; c < headers.Length; c++)
                WriteSaxTextCell(writer, $"{ColName(c + 1)}4", headers[c], ST_TH);
            writer.WriteEndElement();

            uint rIdx = 5;
            foreach (var r in hub.RawData)
            {
                WriteSaxRowStart(writer, rIdx);
                WriteSaxTextCell(writer, $"A{rIdx}", r.OrderCode, ST_TEXT);
                WriteSaxTextCell(writer, $"B{rIdx}", r.OrderDate == default ? "" : r.OrderDate.ToString("dd/MM/yyyy"), ST_DATE);
                WriteSaxTextCell(writer, $"C{rIdx}", r.CustomerName, ST_TEXT);
                WriteSaxTextCell(writer, $"D{rIdx}", r.StoreName, ST_TEXT);
                WriteSaxTextCell(writer, $"E{rIdx}", r.Sku, ST_TEXT);
                WriteSaxTextCell(writer, $"F{rIdx}", r.ProductName, ST_TEXT);
                WriteSaxNumCell(writer, $"G{rIdx}", r.Quantity, ST_NUM);
                WriteSaxNumCell(writer, $"H{rIdx}", r.UnitPrice, ST_CURRENCY);
                WriteSaxNumCell(writer, $"I{rIdx}", r.DiscountAmount, ST_CURRENCY);
                WriteSaxNumCell(writer, $"J{rIdx}", r.UnitCost, ST_CURRENCY);
                WriteSaxTextCell(writer, $"K{rIdx}", r.Status, ST_TEXT);
                WriteSaxTextCell(writer, $"L{rIdx}", r.Channel, ST_TEXT);
                writer.WriteEndElement();
                rIdx++;
            }
            writer.WriteEndElement(); // SheetData

            if (hub.RawData.Count > 0)
                writer.WriteElement(new AutoFilter { Reference = $"A4:L{rIdx - 1}" });

            writer.WriteStartElement(new MergeCells());
            writer.WriteElement(new MergeCell { Reference = "A1:L1" });
            writer.WriteElement(new MergeCell { Reference = "A2:L2" });
            writer.WriteEndElement();

            writer.WriteEndElement(); // Worksheet
        }
        sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "RAW DATA" });
    }

    private static int BuildCleanedDataSheet(WorkbookPart wbPart, Sheets sheets, ref uint sheetId, DataAnalystHub hub)
    {
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        string[] headers = {
            "Mã đơn", "Ngày đơn", "Khách hàng", "Chi nhánh", "SKU", "Tên sản phẩm",
            "Số lượng", "Đơn giá", "Chiết khấu", "Giá vốn ĐV",
            "Doanh thu gộp", "Doanh thu thuần", "Tổng giá vốn", "Lợi nhuận gộp",
            "Biên LN %", "Trạng thái", "Kênh", "Ghi chú làm sạch"
        };

        uint rIdx = 5;
        int lastRow = 5;

        var tablePart = wsPart.AddNewPart<TableDefinitionPart>();
        string tableRelId = wsPart.GetIdOfPart(tablePart);

        using (var writer = OpenXmlWriter.Create(wsPart))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement(new Worksheet());
            WriteSaxColumns(writer, 14, 13, 20, 22, 13, 26, 10, 15, 15, 15, 17, 17, 17, 17, 12, 14, 13, 20);

            writer.WriteStartElement(new SheetData());
            WriteSaxRowStart(writer, 1);
            WriteSaxTextCell(writer, "A1", "CLEANED DATA — DỮ LIỆU ĐÃ LÀM SẠCH & CHUẨN HÓA BI", ST_TITLE);
            writer.WriteEndElement();

            WriteSaxRowStart(writer, 2);
            WriteSaxTextCell(writer, "A2", $"Số bản ghi hợp lệ: {hub.CleanedData.Count} | Kiểm duyệt: {(hub.IsHumanVerified ? "Đã phê duyệt" : "Chờ duyệt")} | Phiên bản #{hub.DataVersion} | Cập nhật: {DateTime.Now:dd/MM/yyyy HH:mm}", ST_SUBTITLE);
            writer.WriteEndElement();

            WriteSaxRowStart(writer, 4, 24);
            for (int c = 0; c < headers.Length; c++)
                WriteSaxTextCell(writer, $"{ColName(c + 1)}4", headers[c], ST_TH);
            writer.WriteEndElement();

            foreach (var r in hub.CleanedData)
            {
                bool ret = r.Status == "RETURNED";
                uint curSt = ret ? ST_NEG_CUR : ST_CURRENCY;
                string sign = ret ? "-" : "";

                WriteSaxRowStart(writer, rIdx);
                WriteSaxTextCell(writer, $"A{rIdx}", r.OrderCode, ST_TEXT);
                WriteSaxTextCell(writer, $"B{rIdx}", r.OrderDate.ToString("yyyy-MM-dd"), ST_DATE);
                WriteSaxTextCell(writer, $"C{rIdx}", r.CustomerName, ST_TEXT);
                WriteSaxTextCell(writer, $"D{rIdx}", r.StoreName, ST_TEXT);
                WriteSaxTextCell(writer, $"E{rIdx}", r.Sku, ST_TEXT);
                WriteSaxTextCell(writer, $"F{rIdx}", r.ProductName, ST_TEXT);
                WriteSaxNumCell(writer, $"G{rIdx}", r.Quantity, ST_NUM);
                WriteSaxNumCell(writer, $"H{rIdx}", r.UnitPrice, ST_CURRENCY);
                WriteSaxNumCell(writer, $"I{rIdx}", r.DiscountAmount, ST_CURRENCY);
                WriteSaxNumCell(writer, $"J{rIdx}", r.UnitCost, ST_CURRENCY);
                WriteSaxFormulaNumCell(writer, $"K{rIdx}", $"{sign}(G{rIdx}*H{rIdx})", r.GrossRevenue, curSt);
                WriteSaxFormulaNumCell(writer, $"L{rIdx}", $"{sign}(G{rIdx}*H{rIdx}-I{rIdx})", r.NetRevenue, curSt);
                WriteSaxFormulaNumCell(writer, $"M{rIdx}", $"{sign}(G{rIdx}*J{rIdx})", r.TotalCost, curSt);
                WriteSaxFormulaNumCell(writer, $"N{rIdx}", $"L{rIdx}-M{rIdx}", r.GrossProfit, r.GrossProfit < 0 ? ST_NEG_CUR : ST_CURRENCY);
                WriteSaxFormulaNumCell(writer, $"O{rIdx}", $"IFERROR(N{rIdx}/L{rIdx},0)", r.NetRevenue != 0 ? r.GrossProfit / r.NetRevenue : 0m, ST_PCT);
                WriteSaxTextCell(writer, $"P{rIdx}", r.Status, ST_TEXT);
                WriteSaxTextCell(writer, $"Q{rIdx}", r.Channel, ST_TEXT);
                WriteSaxTextCell(writer, $"R{rIdx}", r.CleanRemark, ST_TEXT);
                writer.WriteEndElement();
                rIdx++;
            }
            writer.WriteEndElement(); // SheetData

            lastRow = Math.Max(5, (int)rIdx - 1);

            writer.WriteStartElement(new MergeCells());
            writer.WriteElement(new MergeCell { Reference = "A1:R1" });
            writer.WriteElement(new MergeCell { Reference = "A2:R2" });
            writer.WriteEndElement();

            writer.WriteStartElement(new TableParts(), new[] { new OpenXmlAttribute("count", null, "1") });
            writer.WriteElement(new TablePart { Id = tableRelId });
            writer.WriteEndElement();

            writer.WriteEndElement(); // Worksheet
        }

        string tableRef = $"A4:R{lastRow}";
        var table = new Table
        {
            Id = 1,
            Name = "CleanedDataTable",
            DisplayName = "CleanedDataTable",
            Reference = tableRef,
            TotalsRowShown = false
        };
        table.Append(new AutoFilter { Reference = tableRef });
        var tableCols = new TableColumns { Count = (uint)headers.Length };
        for (uint i = 0; i < headers.Length; i++)
            tableCols.Append(new TableColumn { Id = i + 1, Name = headers[i] });
        table.Append(tableCols);
        table.Append(new TableStyleInfo { Name = "TableStyleMedium2", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false });
        tablePart.Table = table;
        tablePart.Table.Save();

        sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "CLEANED DATA" });
        return lastRow;
    }

    private static void BuildDictionarySheet(WorkbookPart wbPart, Sheets sheets, ref uint sheetId)
    {
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sd = new SheetData();
        var cols = MakeColumns(20, 42, 16, 14, 18, 36);

        sd.Append(MakeRow(1, MakeTextCell("A1", "DATA DICTIONARY — TỪ ĐIỂN DỮ LIỆU HỆ THỐNG BI", ST_TITLE)));
        sd.Append(MakeRow(2, MakeTextCell("A2", "Mô tả chuẩn cấu trúc các trường dữ liệu trong Sheet CLEANED DATA", ST_SUBTITLE)));
        sd.Append(MakeHeaderRow(4, new[] { "Tên trường", "Ý nghĩa", "Kiểu dữ liệu", "Đơn vị tính", "Định dạng", "Vai trò trong phân tích BI" }));

        string[][] rows = {
            new[] { "Mã đơn", "Mã định danh duy nhất của đơn hàng", "Văn bản", "Mã", "Text", "Khóa chính kiểm tra trùng lặp & đếm số đơn" },
            new[] { "Ngày đơn", "Ngày phát sinh giao dịch", "Ngày tháng", "Ngày", "YYYY-MM-DD", "Phân tích xu hướng tăng trưởng theo thời gian" },
            new[] { "Khách hàng", "Tên khách hàng hoặc tổ chức mua hàng", "Văn bản", "Tên", "Text", "Phân tích khách hàng & giá trị đơn" },
            new[] { "Chi nhánh", "Cửa hàng / cơ sở thực hiện giao dịch", "Văn bản", "Chi nhánh", "Text", "Phân tích so sánh hiệu quả chi nhánh" },
            new[] { "SKU", "Mã định danh sản phẩm", "Văn bản", "Mã", "Text", "Kết hợp mã đơn để khử trùng & nhóm SKU" },
            new[] { "Tên sản phẩm", "Tên hàng hóa / dịch vụ bán ra", "Văn bản", "Sản phẩm", "Text", "Xếp hạng Top sản phẩm & hiệu quả sản phẩm" },
            new[] { "Số lượng", "Số đơn vị sản phẩm trong dòng giao dịch", "Số nguyên", "Đơn vị", "#,##0", "Tính sản lượng bán & doanh thu gộp" },
            new[] { "Đơn giá", "Giá bán niêm yết trên 1 đơn vị", "Số tiền", "VNĐ", "#,##0 \"đ\"", "Phân tích chính sách giá bán" },
            new[] { "Chiết khấu", "Số tiền giảm giá áp dụng cho dòng đơn", "Số tiền", "VNĐ", "#,##0 \"đ\"", "Đánh giá tỷ lệ & tác động chiết khấu" },
            new[] { "Giá vốn ĐV", "Giá vốn hàng bán trên 1 đơn vị", "Số tiền", "VNĐ", "#,##0 \"đ\"", "Tính tổng giá vốn & lợi nhuận gộp" },
            new[] { "Doanh thu gộp", "Giá trị bán trước chiết khấu (âm nếu hoàn)", "Công thức", "VNĐ", "#,##0 \"đ\"", "Cơ sở tính tỷ lệ chiết khấu" },
            new[] { "Doanh thu thuần", "Doanh thu thực nhận sau chiết khấu & hoàn trả", "Công thức", "VNĐ", "#,##0 \"đ\"", "KPI cốt lõi đánh giá quy mô bán hàng" },
            new[] { "Tổng giá vốn", "Chi phí giá vốn tương ứng sản lượng bán", "Công thức", "VNĐ", "#,##0 \"đ\"", "KPI cấu thành giá thành & lợi nhuận" },
            new[] { "Lợi nhuận gộp", "Chênh lệch giữa Doanh thu thuần và Giá vốn", "Công thức", "VNĐ", "#,##0 \"đ\"", "KPI đo lường hiệu quả sinh lời" },
            new[] { "Biên LN %", "Tỷ suất Lợi nhuận gộp / Doanh thu thuần", "Công thức", "%", "0.00%", "Đánh giá chất lượng lợi nhuận theo nhóm" },
            new[] { "Trạng thái", "Tình trạng đơn (COMPLETED / RETURNED)", "Văn bản", "Trạng thái", "Text", "Phân tách đơn hoàn trả & tính tỷ lệ hoàn" },
            new[] { "Kênh", "Kênh bán hàng (Store, Online, B2B...)", "Văn bản", "Kênh", "Text", "Phân tích cơ cấu đóng góp theo kênh" },
            new[] { "Ghi chú làm sạch", "Trạng thái xử lý từ bước Làm sạch", "Văn bản", "Nhãn", "Text", "Truy vết minh bạch dữ liệu ngoại lệ" }
        };

        uint rIdx = 5;
        foreach (var item in rows)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", item[0], ST_TEXT_BOLD),
                MakeTextCell($"B{rIdx}", item[1], ST_TEXT),
                MakeTextCell($"C{rIdx}", item[2], ST_TEXT),
                MakeTextCell($"D{rIdx}", item[3], ST_TEXT),
                MakeTextCell($"E{rIdx}", item[4], ST_TEXT),
                MakeTextCell($"F{rIdx}", item[5], ST_TEXT)
            );
            sd.Append(row);
            rIdx++;
        }

        var merges = new MergeCells();
        merges.Append(new MergeCell { Reference = "A1:F1" }, new MergeCell { Reference = "A2:F2" });
        var ws = new Worksheet();
        ws.Append(cols, sd, new AutoFilter { Reference = $"A4:F{rIdx - 1}" }, merges);
        wsPart.Worksheet = ws;
        wsPart.Worksheet.Save();
        sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "DATA DICTIONARY" });
    }

    private static void BuildCleaningLogSheet(WorkbookPart wbPart, Sheets sheets, ref uint sheetId, DataAnalystHub hub)
    {
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sd = new SheetData();
        var cols = MakeColumns(28, 22, 26, 44, 24);

        sd.Append(MakeRow(1, MakeTextCell("A1", "CLEANING LOG — NHẬT KÝ & THỐNG KÊ LÀM SẠCH DỮ LIỆU", ST_TITLE)));
        sd.Append(MakeRow(2, MakeTextCell("A2", $"Nguồn dữ liệu: {hub.SourceName} | Thời điểm: {DateTime.Now:dd/MM/yyyy HH:mm}", ST_SUBTITLE)));

        sd.Append(MakeRow(4, MakeTextCell("A4", "1. BẢNG TỔNG HỢP KẾT QUẢ LÀM SẠCH", ST_SECTION)));
        sd.Append(MakeHeaderRow(5, new[] { "Chỉ tiêu kiểm soát", "Số lượng / Giá trị", "Quy tắc xử lý", "Ý nghĩa nghiệp vụ", "Trạng thái" }));

        var st = hub.LastStats;
        var summaryRows = new (string label, decimal val, uint style, string rule, string meaning, string status)[]
        {
            ("Tổng số bản ghi đầu vào", st.RawCount, ST_NUM, "Đọc toàn bộ từ tệp nguồn", "Quy mô dữ liệu thô ban đầu", "Đã tiếp nhận"),
            ("Bản ghi trùng lặp bị loại", st.DuplicatesRemoved, ST_NUM, "Khóa (Mã đơn + SKU + Sản phẩm)", "Ngăn chặn tính trùng doanh thu", "Đã loại bỏ"),
            ("Đơn hàng hủy / lỗi POS", st.CancelledExcluded, ST_NUM, "Trạng thái CANCELLED / FAILED", "Không ghi nhận doanh thu ảo", "Đã loại bỏ"),
            ("Đơn hàng hoàn trả", st.ReturnsAdjusted, ST_NUM, "Trạng thái RETURNED → ghi âm", "Khấu trừ doanh thu & hoàn giá vốn", "Đã đối soát"),
            ("Dữ liệu thiếu / sai định dạng", st.MissingOrInvalid, ST_NUM, "SL <= 0, Đơn giá <= 0 hoặc thiếu mã", "Tách sang bảng ngoại lệ kiểm tra", st.MissingOrInvalid > 0 ? "Ngoại lệ" : "Không có"),
            ("Dòng thiếu giá vốn đơn vị", st.ImputedCount, ST_NUM, "Để trống (0 đ)", "Ghi nhận trạng thái thiếu trường giá vốn", "Đã ghi nhận"),
            ("Bản ghi hợp lệ sau làm sạch", st.CleanedCount, ST_NUM, "Đưa vào Sheet CLEANED DATA", "Dữ liệu chuẩn phục vụ BI", hub.IsHumanVerified ? "Đã phê duyệt" : "Chờ duyệt"),
            ("Giá trị sai lệch đã ngăn chặn", st.PreventedOvercountValue, ST_CURRENCY, "Tổng giá trị đơn trùng + hủy", "Số tiền bảo vệ báo cáo khỏi sai lệch", "Đã kiểm soát")
        };

        uint rIdx = 6;
        foreach (var sr in summaryRows)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", sr.label, ST_TEXT_BOLD),
                MakeNumCell($"B{rIdx}", sr.val, sr.style),
                MakeTextCell($"C{rIdx}", sr.rule, ST_TEXT),
                MakeTextCell($"D{rIdx}", sr.meaning, ST_TEXT),
                MakeTextCell($"E{rIdx}", sr.status, ST_TEXT)
            );
            sd.Append(row);
            rIdx++;
        }

        rIdx += 1;
        sd.Append(MakeRow(rIdx, MakeTextCell($"A{rIdx}", "2. CHI TIẾT NHẬT KÝ & BẢN GHI NGOẠI LỆ", ST_SECTION)));
        rIdx++;
        sd.Append(MakeHeaderRow(rIdx, new[] { "STT / Loại", "Mã đơn / Dòng", "Lý do phát hiện", "Hành động xử lý", "Chi tiết nhật ký" }));
        rIdx++;

        int idx = 1;
        foreach (var ex in hub.Exceptions)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", $"Ngoại lệ #{idx} ({ex.ExceptionType})", ST_TEXT_BOLD),
                MakeTextCell($"B{rIdx}", ex.Record.OrderCode, ST_TEXT),
                MakeTextCell($"C{rIdx}", ex.Reason, ST_TEXT),
                MakeTextCell($"D{rIdx}", ex.ActionTaken, ST_TEXT),
                MakeTextCell($"E{rIdx}", $"{ex.Record.ProductName} · SL {ex.Record.Quantity} · Giá {ex.Record.UnitPrice:N0}", ST_TEXT)
            );
            sd.Append(row);
            rIdx++; idx++;
        }

        foreach (var log in hub.CleaningLogs)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", "System Log", ST_TEXT),
                MakeTextCell($"B{rIdx}", "-", ST_TEXT),
                MakeTextCell($"C{rIdx}", "Tự động", ST_TEXT),
                MakeTextCell($"D{rIdx}", "Ghi nhận", ST_TEXT),
                MakeTextCell($"E{rIdx}", log, ST_TEXT)
            );
            sd.Append(row);
            rIdx++;
        }

        var merges = new MergeCells();
        merges.Append(new MergeCell { Reference = "A1:E1" }, new MergeCell { Reference = "A2:E2" });
        var ws = new Worksheet();
        ws.Append(cols, sd, merges);
        wsPart.Worksheet = ws;
        wsPart.Worksheet.Save();
        sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "CLEANING LOG" });
    }

    private class AnalysisRanges
    {
        public string KpiRange { get; set; } = "";
        public int StoreHeaderRow { get; set; }
        public int StoreStartRow { get; set; }
        public int StoreEndRow { get; set; }
        public int ProdHeaderRow { get; set; }
        public int ProdStartRow { get; set; }
        public int ProdEndRow { get; set; }
        public int TrendHeaderRow { get; set; }
        public int TrendStartRow { get; set; }
        public int TrendEndRow { get; set; }
        public int ChannelHeaderRow { get; set; }
        public int ChannelStartRow { get; set; }
        public int ChannelEndRow { get; set; }
    }

    private static AnalysisRanges BuildAnalysisSheet(WorkbookPart wbPart, Sheets sheets, ref uint sheetId, DataAnalystHub hub, DeepAnalysisResult analysis, int cleanLastRow)
    {
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sd = new SheetData();
        var cols = MakeColumns(30, 18, 18, 18, 16, 16, 22);
        var ranges = new AnalysisRanges();

        sd.Append(MakeRow(1, MakeTextCell("A1", "ANALYSIS — BẢNG TỔNG HỢP PHÂN TÍCH KINH DOANH (CÔNG THỨC ĐỘNG)", ST_TITLE)));
        sd.Append(MakeRow(2, MakeTextCell("A2", "Mọi chỉ số dưới đây sử dụng công thức tham chiếu trực tiếp sang Sheet CLEANED DATA.", ST_SUBTITLE)));

        string cd = "'CLEANED DATA'";
        string rNet = $"{cd}!$L$5:$L${cleanLastRow}";
        string rGross = $"{cd}!$K$5:$K${cleanLastRow}";
        string rCost = $"{cd}!$M$5:$M${cleanLastRow}";
        string rProfit = $"{cd}!$N$5:$N${cleanLastRow}";
        string rDisc = $"{cd}!$I$5:$I${cleanLastRow}";
        string rStatus = $"{cd}!$P$5:$P${cleanLastRow}";
        string rStore = $"{cd}!$D$5:$D${cleanLastRow}";
        string rProd = $"{cd}!$F$5:$F${cleanLastRow}";
        string rQty = $"{cd}!$G$5:$G${cleanLastRow}";
        string rDate = $"{cd}!$B$5:$B${cleanLastRow}";
        string rChan = $"{cd}!$Q$5:$Q${cleanLastRow}";

        // 1. Tổng quan doanh thu, lợi nhuận, giá bán & chiết khấu
        sd.Append(MakeRow(4, MakeTextCell("A4", "1. CHỈ SỐ TỔNG QUAN DOANH THU, LỢI NHUẬN & CHIẾT KHẤU", ST_SECTION)));
        sd.Append(MakeHeaderRow(5, new[] { "Chỉ số BI", "Giá trị", "Đơn vị", "Diễn giải" }));

        var m = analysis.OverallMetrics;
        var p = analysis.PricingAnalysis;
        var kpis = new (string name, string formula, double cached, uint style, string unit, string desc)[]
        {
            ("Tổng doanh thu gộp", $"SUMIF({rGross},\">0\")", (double)m.TotalGrossRevenue, ST_CURRENCY, "VNĐ", "Doanh thu trước chiết khấu"),
            ("Tổng chiết khấu", $"SUMIF({rStatus},\"<>RETURNED\",{rDisc})", (double)m.TotalDiscountAmount, ST_CURRENCY, "VNĐ", "Tổng giảm giá trên đơn bán"),
            ("Tổng doanh thu thuần", $"SUM({rNet})", (double)m.TotalNetRevenue, ST_CURRENCY, "VNĐ", "Doanh thu thực nhận sau hoàn trả"),
            ("Tổng giá vốn (COGS)", $"SUM({rCost})", (double)m.TotalCost, ST_CURRENCY, "VNĐ", "Chi phí giá vốn hàng bán"),
            ("Tổng lợi nhuận gộp", $"SUM({rProfit})", (double)m.TotalGrossProfit, ST_CURRENCY, "VNĐ", "Doanh thu thuần - Giá vốn"),
            ("Biên lợi nhuận trung bình", "IFERROR(B10/B8,0)", (double)(m.GrossMarginPercent / 100m), ST_PCT, "%", "Lợi nhuận gộp / Doanh thu thuần"),
            ("Tổng số đơn hàng hợp lệ", $"COUNTA({cd}!$A$5:$A${cleanLastRow})", m.TotalOrders, ST_NUM, "Đơn", "Số bản ghi sạch đã hạch toán"),
            ("Giá trị đơn hàng trung bình", $"IFERROR(AVERAGEIFS({rNet},{rStatus},\"<>RETURNED\"),0)", (double)p.AverageOrderValue, ST_CURRENCY, "VNĐ", "Doanh thu thuần TB mỗi đơn bán"),
            ("Tỷ lệ chiết khấu trung bình", "IFERROR(B7/B6,0)", (double)(p.AverageDiscountRate / 100m), ST_PCT, "%", "Tổng chiết khấu / Doanh thu gộp"),
            ("Tỷ lệ hoàn trả", $"IFERROR(COUNTIF({rStatus},\"RETURNED\")/B12,0)", (double)(m.ReturnRatePercent / 100m), ST_PCT, "%", "Số đơn hoàn / Tổng đơn hợp lệ")
        };

        uint rIdx = 6;
        foreach (var k in kpis)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", k.name, ST_TEXT_BOLD),
                MakeFormulaNumCell($"B{rIdx}", k.formula, (decimal)k.cached, k.style),
                MakeTextCell($"C{rIdx}", k.unit, ST_TEXT),
                MakeTextCell($"D{rIdx}", k.desc, ST_TEXT)
            );
            sd.Append(row);
            rIdx++;
        }

        // 2. Phân tích theo Chi nhánh
        rIdx++;
        sd.Append(MakeRow(rIdx, MakeTextCell($"A{rIdx}", "2. PHÂN TÍCH HIỆU QUẢ CHI NHÁNH", ST_SECTION)));
        rIdx++;
        ranges.StoreHeaderRow = (int)rIdx;
        sd.Append(MakeHeaderRow(rIdx, new[] { "Chi nhánh", "Số đơn", "Doanh thu thuần", "Lợi nhuận gộp", "Biên LN %", "Tỷ lệ hoàn %", "Đánh giá" }));
        rIdx++;
        ranges.StoreStartRow = (int)rIdx;
        foreach (var s in analysis.StorePerformance)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", s.StoreName, ST_TEXT_BOLD),
                MakeFormulaNumCell($"B{rIdx}", $"COUNTIF({rStore},A{rIdx})", s.OrderCount, ST_NUM),
                MakeFormulaNumCell($"C{rIdx}", $"SUMIF({rStore},A{rIdx},{rNet})", s.NetRevenue, ST_CURRENCY),
                MakeFormulaNumCell($"D{rIdx}", $"SUMIF({rStore},A{rIdx},{rProfit})", s.GrossProfit, ST_CURRENCY),
                MakeFormulaNumCell($"E{rIdx}", $"IFERROR(D{rIdx}/C{rIdx},0)", s.MarginPercent / 100m, ST_PCT),
                MakeFormulaNumCell($"F{rIdx}", $"IFERROR(COUNTIFS({rStore},A{rIdx},{rStatus},\"RETURNED\")/B{rIdx},0)", s.ReturnRatePercent / 100m, ST_PCT),
                MakeTextCell($"G{rIdx}", s.StatusAssessment, ST_TEXT)
            );
            sd.Append(row);
            rIdx++;
        }
        ranges.StoreEndRow = Math.Max(ranges.StoreStartRow, (int)rIdx - 1);

        // 3. Phân tích theo Sản phẩm
        rIdx++;
        sd.Append(MakeRow(rIdx, MakeTextCell($"A{rIdx}", "3. PHÂN TÍCH HIỆU QUẢ SẢN PHẨM (TOP DOANH THU)", ST_SECTION)));
        rIdx++;
        ranges.ProdHeaderRow = (int)rIdx;
        sd.Append(MakeHeaderRow(rIdx, new[] { "Tên sản phẩm", "Sản lượng", "Doanh thu thuần", "Lợi nhuận gộp", "Biên LN %", "Chiết khấu", "Hiệu quả" }));
        rIdx++;
        ranges.ProdStartRow = (int)rIdx;
        foreach (var pr in analysis.ProductPerformance.Take(12))
        {
            string eff = pr.MarginPercent < 10 ? "Biên LN thấp" : pr.NetRevenue <= 0 ? "Hiệu quả kém" : "Hiệu quả tốt";
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", pr.ProductName, ST_TEXT_BOLD),
                MakeFormulaNumCell($"B{rIdx}", $"SUMIFS({rQty},{rProd},A{rIdx},{rStatus},\"<>RETURNED\")-SUMIFS({rQty},{rProd},A{rIdx},{rStatus},\"RETURNED\")", pr.QuantitySold, ST_NUM),
                MakeFormulaNumCell($"C{rIdx}", $"SUMIF({rProd},A{rIdx},{rNet})", pr.NetRevenue, ST_CURRENCY),
                MakeFormulaNumCell($"D{rIdx}", $"SUMIF({rProd},A{rIdx},{rProfit})", pr.GrossProfit, ST_CURRENCY),
                MakeFormulaNumCell($"E{rIdx}", $"IFERROR(D{rIdx}/C{rIdx},0)", pr.MarginPercent / 100m, ST_PCT),
                MakeFormulaNumCell($"F{rIdx}", $"SUMIFS({rDisc},{rProd},A{rIdx},{rStatus},\"<>RETURNED\")", pr.DiscountAmount, ST_CURRENCY),
                MakeTextCell($"G{rIdx}", eff, ST_TEXT)
            );
            sd.Append(row);
            rIdx++;
        }
        ranges.ProdEndRow = Math.Max(ranges.ProdStartRow, (int)rIdx - 1);

        // 4. Xu hướng theo thời gian
        rIdx++;
        sd.Append(MakeRow(rIdx, MakeTextCell($"A{rIdx}", "4. XU HƯỚNG DOANH THU & LỢI NHUẬN THEO THỜI GIAN", ST_SECTION)));
        rIdx++;
        ranges.TrendHeaderRow = (int)rIdx;
        sd.Append(MakeHeaderRow(rIdx, new[] { "Ngày", "Số đơn", "Doanh thu thuần", "Lợi nhuận gộp", "Biên LN %" }));
        rIdx++;
        ranges.TrendStartRow = (int)rIdx;
        foreach (var t in analysis.TimeSeries)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", t.Date, ST_DATE),
                MakeFormulaNumCell($"B{rIdx}", $"COUNTIF({rDate},A{rIdx})", t.Orders, ST_NUM),
                MakeFormulaNumCell($"C{rIdx}", $"SUMIF({rDate},A{rIdx},{rNet})", t.NetRevenue, ST_CURRENCY),
                MakeFormulaNumCell($"D{rIdx}", $"SUMIF({rDate},A{rIdx},{rProfit})", t.GrossProfit, ST_CURRENCY),
                MakeFormulaNumCell($"E{rIdx}", $"IFERROR(D{rIdx}/C{rIdx},0)", t.NetRevenue != 0 ? t.GrossProfit / t.NetRevenue : 0m, ST_PCT)
            );
            sd.Append(row);
            rIdx++;
        }
        ranges.TrendEndRow = Math.Max(ranges.TrendStartRow, (int)rIdx - 1);

        // 5. Cơ cấu theo Kênh
        rIdx++;
        sd.Append(MakeRow(rIdx, MakeTextCell($"A{rIdx}", "5. CƠ CẤU DOANH THU THEO KÊNH BÁN HÀNG", ST_SECTION)));
        rIdx++;
        ranges.ChannelHeaderRow = (int)rIdx;
        sd.Append(MakeHeaderRow(rIdx, new[] { "Kênh bán", "Số đơn", "Doanh thu thuần", "Tỷ trọng %" }));
        rIdx++;
        ranges.ChannelStartRow = (int)rIdx;
        foreach (var ch in analysis.ChannelBreakdown)
        {
            var row = new Row { RowIndex = rIdx };
            row.Append(
                MakeTextCell($"A{rIdx}", ch.ChannelName, ST_TEXT_BOLD),
                MakeFormulaNumCell($"B{rIdx}", $"COUNTIF({rChan},A{rIdx})", ch.OrderCount, ST_NUM),
                MakeFormulaNumCell($"C{rIdx}", $"SUMIF({rChan},A{rIdx},{rNet})", ch.NetRevenue, ST_CURRENCY),
                MakeFormulaNumCell($"D{rIdx}", $"IFERROR(C{rIdx}/$B$8,0)", ch.RevenueSharePercent / 100m, ST_PCT)
            );
            sd.Append(row);
            rIdx++;
        }
        ranges.ChannelEndRow = Math.Max(ranges.ChannelStartRow, (int)rIdx - 1);

        var merges = new MergeCells();
        merges.Append(new MergeCell { Reference = "A1:G1" }, new MergeCell { Reference = "A2:G2" });
        var ws = new Worksheet();
        ws.Append(cols, sd, merges);
        wsPart.Worksheet = ws;
        wsPart.Worksheet.Save();
        sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "ANALYSIS" });
        return ranges;
    }

    private static void BuildDashboardSheet(WorkbookPart wbPart, Sheets sheets, ref uint sheetId, DataAnalystHub hub, DeepAnalysisResult analysis, AnalysisRanges ranges)
    {
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sd = new SheetData();
        var cols = MakeColumns(18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18);

        sd.Append(MakeRow(1, MakeTextCell("A1", "NGỌC LINH DATA STUDIO — DASHBOARD ĐIỀU HÀNH BI", ST_TITLE)));
        sd.Append(MakeRow(2, MakeTextCell("A2", $"Nguồn: {hub.SourceName} | Cập nhật: {DateTime.Now:dd/MM/yyyy HH:mm} | Dữ liệu sạch: {hub.CleanedData.Count} bản ghi", ST_SUBTITLE)));

        var m = analysis.OverallMetrics;
        var p = analysis.PricingAnalysis;

        // KPI Cards (Rows 4-5)
        var r4 = new Row { RowIndex = 4 };
        r4.Append(
            MakeTextCell("A4", "TỔNG DOANH THU THUẦN", ST_KPI_LABEL), MakeTextCell("B4", "", ST_KPI_LABEL),
            MakeTextCell("C4", "TỔNG LỢI NHUẬN GỘP", ST_KPI_LABEL), MakeTextCell("D4", "", ST_KPI_LABEL),
            MakeTextCell("E4", "BIÊN LỢI NHUẬN TB", ST_KPI_LABEL), MakeTextCell("F4", "", ST_KPI_LABEL),
            MakeTextCell("G4", "ĐƠN HÀNG HỢP LỆ", ST_KPI_LABEL), MakeTextCell("H4", "", ST_KPI_LABEL),
            MakeTextCell("I4", "GIÁ TRỊ ĐƠN TRUNG BÌNH", ST_KPI_LABEL), MakeTextCell("J4", "", ST_KPI_LABEL)
        );
        sd.Append(r4);

        var r5 = new Row { RowIndex = 5, Height = 28, CustomHeight = true };
        r5.Append(
            MakeFormulaNumCell("A5", "ANALYSIS!B8", m.TotalNetRevenue, ST_KPI_VAL_CUR), MakeTextCell("B5", "", ST_KPI_VAL_CUR),
            MakeFormulaNumCell("C5", "ANALYSIS!B10", m.TotalGrossProfit, ST_KPI_VAL_CUR), MakeTextCell("D5", "", ST_KPI_VAL_CUR),
            MakeFormulaNumCell("E5", "ANALYSIS!B11", m.GrossMarginPercent / 100m, ST_KPI_VAL_PCT), MakeTextCell("F5", "", ST_KPI_VAL_PCT),
            MakeFormulaNumCell("G5", "ANALYSIS!B12", m.TotalOrders, ST_KPI_VAL_NUM), MakeTextCell("H5", "", ST_KPI_VAL_NUM),
            MakeFormulaNumCell("I5", "ANALYSIS!B13", p.AverageOrderValue, ST_KPI_VAL_CUR), MakeTextCell("J5", "", ST_KPI_VAL_CUR)
        );
        sd.Append(r5);

        // Insight tổng quan (Rows 7-12)
        sd.Append(MakeRow(7, MakeTextCell("A7", "INSIGHT TỔNG QUAN TỪ DỮ LIỆU THỰC TẾ", ST_SECTION)));
        var bestStore = analysis.StorePerformance.FirstOrDefault();
        var bestProd = analysis.ProductPerformance.OrderByDescending(x => x.GrossProfit).FirstOrDefault();
        var lowMarginProd = analysis.ProductPerformance.Where(x => x.MarginPercent < 12).FirstOrDefault();

        var insights = new (string k, string v)[]
        {
            ("Chi nhánh doanh thu cao nhất", bestStore != null ? $"{bestStore.StoreName}: {bestStore.NetRevenue:N0} đ (Biên LN {bestStore.MarginPercent}%)" : "Chưa có dữ liệu chi nhánh"),
            ("Sản phẩm lợi nhuận tốt nhất", bestProd != null ? $"{bestProd.ProductName}: LN gộp {bestProd.GrossProfit:N0} đ (Biên LN {bestProd.MarginPercent}%)" : "Chưa có dữ liệu sản phẩm"),
            ("Sản phẩm biên lợi nhuận thấp", lowMarginProd != null ? $"{lowMarginProd.ProductName}: Biên LN {lowMarginProd.MarginPercent}%" : "Không có sản phẩm nào biên LN < 12%"),
            ("Xu hướng & Chất lượng đơn", $"Tăng trưởng theo ngày: {analysis.GrowthRatePercent}% | Tỷ lệ hoàn trả: {m.ReturnRatePercent}% | Sai lệch đã ngăn chặn: {m.PreventedOvercountValue:N0} đ")
        };

        uint iRow = 8;
        foreach (var ins in insights)
        {
            var row = new Row { RowIndex = iRow };
            row.Append(MakeTextCell($"A{iRow}", ins.k, ST_TEXT_BOLD), MakeTextCell($"C{iRow}", ins.v, ST_TEXT));
            sd.Append(row);
            iRow++;
        }

        sd.Append(MakeRow(13, MakeTextCell("A13", "Gợi ý lọc dữ liệu: Sử dụng AutoFilter trên Sheet 'CLEANED DATA' hoặc bảng tổng hợp tại 'ANALYSIS' theo Thời gian, Chi nhánh, Sản phẩm, Kênh.", ST_SUBTITLE)));

        var merges = new MergeCells();
        merges.Append(
            new MergeCell { Reference = "A1:J1" },
            new MergeCell { Reference = "A2:J2" },
            new MergeCell { Reference = "A4:B4" }, new MergeCell { Reference = "A5:B5" },
            new MergeCell { Reference = "C4:D4" }, new MergeCell { Reference = "C5:D5" },
            new MergeCell { Reference = "E4:F4" }, new MergeCell { Reference = "E5:F5" },
            new MergeCell { Reference = "G4:H4" }, new MergeCell { Reference = "G5:H5" },
            new MergeCell { Reference = "I4:J4" }, new MergeCell { Reference = "I5:J5" },
            new MergeCell { Reference = "A7:J7" },
            new MergeCell { Reference = "A8:B8" }, new MergeCell { Reference = "C8:J8" },
            new MergeCell { Reference = "A9:B9" }, new MergeCell { Reference = "C9:J9" },
            new MergeCell { Reference = "A10:B10" }, new MergeCell { Reference = "C10:J10" },
            new MergeCell { Reference = "A11:B11" }, new MergeCell { Reference = "C11:J11" },
            new MergeCell { Reference = "A13:J13" }
        );

        // Add Native Excel Charts via DrawingsPart
        var drawingsPart = wsPart.AddNewPart<DrawingsPart>();
        var wsDrawing = new Xdr.WorksheetDrawing();

        // Chart 1: Line chart - Xu hướng doanh thu & lợi nhuận theo thời gian
        AddChartToDrawing(drawingsPart, wsDrawing, "Xu hướng Doanh thu & Lợi nhuận theo thời gian", ChartKind.Line,
            $"ANALYSIS!$A${ranges.TrendStartRow}:$A${ranges.TrendEndRow}",
            new[] { ("Doanh thu thuần", $"ANALYSIS!$C${ranges.TrendStartRow}:$C${ranges.TrendEndRow}"), ("Lợi nhuận gộp", $"ANALYSIS!$D${ranges.TrendStartRow}:$D${ranges.TrendEndRow}") },
            0, 14, 5, 29, 100);

        // Chart 2: Column chart - So sánh doanh thu giữa các chi nhánh
        AddChartToDrawing(drawingsPart, wsDrawing, "So sánh Doanh thu & Lợi nhuận theo Chi nhánh", ChartKind.Column,
            $"ANALYSIS!$A${ranges.StoreStartRow}:$A${ranges.StoreEndRow}",
            new[] { ("Doanh thu thuần", $"ANALYSIS!$C${ranges.StoreStartRow}:$C${ranges.StoreEndRow}"), ("Lợi nhuận gộp", $"ANALYSIS!$D${ranges.StoreStartRow}:$D${ranges.StoreEndRow}") },
            5, 14, 10, 29, 200);

        // Chart 3: Bar chart - Top sản phẩm theo doanh thu
        AddChartToDrawing(drawingsPart, wsDrawing, "Top Sản phẩm theo Doanh thu thuần", ChartKind.Bar,
            $"ANALYSIS!$A${ranges.ProdStartRow}:$A${ranges.ProdEndRow}",
            new[] { ("Doanh thu thuần", $"ANALYSIS!$C${ranges.ProdStartRow}:$C${ranges.ProdEndRow}") },
            0, 30, 5, 45, 300);

        // Chart 4: Pie chart - Cơ cấu doanh thu theo Kênh bán
        AddChartToDrawing(drawingsPart, wsDrawing, "Cơ cấu Doanh thu theo Kênh bán", ChartKind.Pie,
            $"ANALYSIS!$A${ranges.ChannelStartRow}:$A${ranges.ChannelEndRow}",
            new[] { ("Doanh thu thuần", $"ANALYSIS!$C${ranges.ChannelStartRow}:$C${ranges.ChannelEndRow}") },
            5, 30, 10, 45, 400);

        drawingsPart.WorksheetDrawing = wsDrawing;
        drawingsPart.WorksheetDrawing.Save();

        var ws = new Worksheet();
        ws.Append(cols, sd, merges, new Drawing { Id = wsPart.GetIdOfPart(drawingsPart) });
        wsPart.Worksheet = ws;
        wsPart.Worksheet.Save();
        sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId++, Name = "DASHBOARD" });
    }

    private enum ChartKind { Line, Column, Bar, Pie }

    private static void AddChartToDrawing(DrawingsPart drawingsPart, Xdr.WorksheetDrawing wsDrawing, string title, ChartKind kind,
        string catFormula, (string name, string valFormula)[] seriesList, int col1, int row1, int col2, int row2, uint axisBaseId)
    {
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        var chartSpace = new ChartSpace();
        chartSpace.Append(new EditingLanguage { Val = "vi-VN" });
        var chart = chartSpace.AppendChild(new C.Chart());

        var titleEl = new C.Title();
        titleEl.Append(new C.ChartText(new C.RichText(
            new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.Run(new A.RunProperties { Bold = true, FontSize = 1100 }, new A.Text(title)))
        )));
        titleEl.Append(new Overlay { Val = false });
        chart.Append(titleEl);
        chart.Append(new AutoTitleDeleted { Val = false });

        var plotArea = chart.AppendChild(new PlotArea());
        plotArea.Append(new Layout());

        uint axCat = axisBaseId + 1;
        uint axVal = axisBaseId + 2;

        if (kind == ChartKind.Pie)
        {
            var pie = plotArea.AppendChild(new PieChart(new VaryColors { Val = true }));
            for (uint i = 0; i < seriesList.Length; i++)
            {
                var ser = pie.AppendChild(new PieChartSeries(
                    new C.Index { Val = i }, new Order { Val = i },
                    new SeriesText(new NumericValue { Text = seriesList[i].name })
                ));
                ser.Append(new CategoryAxisData(new StringReference(new C.Formula { Text = catFormula })));
                ser.Append(new C.Values(new NumberReference(new C.Formula { Text = seriesList[i].valFormula })));
            }
        }
        else if (kind == ChartKind.Line)
        {
            var line = plotArea.AppendChild(new LineChart(new Grouping { Val = GroupingValues.Standard }, new VaryColors { Val = false }));
            for (uint i = 0; i < seriesList.Length; i++)
            {
                var ser = line.AppendChild(new LineChartSeries(
                    new C.Index { Val = i }, new Order { Val = i },
                    new SeriesText(new NumericValue { Text = seriesList[i].name }),
                    new Marker(new Symbol { Val = MarkerStyleValues.Circle })
                ));
                ser.Append(new CategoryAxisData(new StringReference(new C.Formula { Text = catFormula })));
                ser.Append(new C.Values(new NumberReference(new C.Formula { Text = seriesList[i].valFormula })));
            }
            line.Append(new AxisId { Val = axCat }, new AxisId { Val = axVal });
            AppendAxes(plotArea, axCat, axVal, false);
        }
        else
        {
            var barDir = kind == ChartKind.Bar ? BarDirectionValues.Bar : BarDirectionValues.Column;
            var bar = plotArea.AppendChild(new BarChart(new BarDirection { Val = barDir }, new BarGrouping { Val = BarGroupingValues.Clustered }, new VaryColors { Val = false }));
            for (uint i = 0; i < seriesList.Length; i++)
            {
                var ser = bar.AppendChild(new BarChartSeries(
                    new C.Index { Val = i }, new Order { Val = i },
                    new SeriesText(new NumericValue { Text = seriesList[i].name })
                ));
                ser.Append(new CategoryAxisData(new StringReference(new C.Formula { Text = catFormula })));
                ser.Append(new C.Values(new NumberReference(new C.Formula { Text = seriesList[i].valFormula })));
            }
            bar.Append(new AxisId { Val = axCat }, new AxisId { Val = axVal });
            AppendAxes(plotArea, axCat, axVal, kind == ChartKind.Bar);
        }

        chart.Append(new Legend(new LegendPosition { Val = LegendPositionValues.Bottom }, new Overlay { Val = false }));
        chart.Append(new PlotVisibleOnly { Val = true });
        chartSpace.Append(new C.PrintSettings());
        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();

        var anchor = new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(new Xdr.ColumnId(col1.ToString()), new Xdr.ColumnOffset("0"), new Xdr.RowId(row1.ToString()), new Xdr.RowOffset("0")),
            new Xdr.ToMarker(new Xdr.ColumnId(col2.ToString()), new Xdr.ColumnOffset("0"), new Xdr.RowId(row2.ToString()), new Xdr.RowOffset("0")),
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = axisBaseId, Name = title },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()
                ),
                new Xdr.Transform(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 0, Cy = 0 }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
            ),
            new Xdr.ClientData()
        );
        wsDrawing.Append(anchor);
    }

    private static void AppendAxes(PlotArea plotArea, uint axCat, uint axVal, bool isHorizontalBar)
    {
        plotArea.Append(new CategoryAxis(
            new AxisId { Val = axCat },
            new Scaling(new Orientation { Val = C.OrientationValues.MinMax }),
            new Delete { Val = false },
            new AxisPosition { Val = isHorizontalBar ? AxisPositionValues.Left : AxisPositionValues.Bottom },
            new TickLabelPosition { Val = TickLabelPositionValues.NextTo },
            new CrossingAxis { Val = axVal },
            new Crosses { Val = CrossesValues.AutoZero }
        ));
        plotArea.Append(new ValueAxis(
            new AxisId { Val = axVal },
            new Scaling(new Orientation { Val = C.OrientationValues.MinMax }),
            new Delete { Val = false },
            new AxisPosition { Val = isHorizontalBar ? AxisPositionValues.Bottom : AxisPositionValues.Left },
            new MajorGridlines(),
            new C.NumberingFormat { FormatCode = "#,##0", SourceLinked = true },
            new TickLabelPosition { Val = TickLabelPositionValues.NextTo },
            new CrossingAxis { Val = axCat },
            new Crosses { Val = CrossesValues.AutoZero }
        ));
    }

    // ---------------- Helpers ----------------
    private static Columns MakeColumns(params double[] widths)
    {
        var cols = new Columns();
        for (uint i = 0; i < widths.Length; i++)
            cols.Append(new Column { Min = i + 1, Max = i + 1, Width = widths[i], CustomWidth = true });
        return cols;
    }

    private static Row MakeRow(uint idx, params Cell[] cells)
    {
        var r = new Row { RowIndex = idx };
        r.Append(cells);
        return r;
    }

    private static Row MakeHeaderRow(uint idx, string[] headers)
    {
        var r = new Row { RowIndex = idx, Height = 22, CustomHeight = true };
        for (int i = 0; i < headers.Length; i++)
            r.Append(MakeTextCell($"{ColName(i)}{idx}", headers[i], ST_TH));
        return r;
    }

    private static string ColName(int idx)
    {
        string s = "";
        idx++;
        while (idx > 0) { int m = (idx - 1) % 26; s = (char)('A' + m) + s; idx = (idx - m) / 26; }
        return s;
    }

    private static Cell MakeTextCell(string cellRef, string text, uint style) =>
        new() { CellReference = cellRef, DataType = CellValues.InlineString, StyleIndex = style, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(text ?? "")) };

    private static Cell MakeNumCell(string cellRef, decimal val, uint style) =>
        new() { CellReference = cellRef, DataType = CellValues.Number, StyleIndex = style, CellValue = new CellValue(val.ToString(CultureInfo.InvariantCulture)) };

    private static Cell MakeFormulaNumCell(string cellRef, string formula, decimal cachedVal, uint style) =>
        new() { CellReference = cellRef, DataType = CellValues.Number, StyleIndex = style, CellFormula = new CellFormula(formula), CellValue = new CellValue(cachedVal.ToString(CultureInfo.InvariantCulture)) };

    private static Stylesheet BuildStylesheet()
    {
        var fonts = new Fonts(
            new DocumentFormat.OpenXml.Spreadsheet.Font(new FontSize { Val = 10 }, new FontName { Val = "Calibri" }), // 0 default
            new DocumentFormat.OpenXml.Spreadsheet.Font(new Bold(), new FontSize { Val = 14 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF3C4430" }, new FontName { Val = "Calibri" }), // 1 title
            new DocumentFormat.OpenXml.Spreadsheet.Font(new Italic(), new FontSize { Val = 10 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF666666" }, new FontName { Val = "Calibri" }), // 2 subtitle
            new DocumentFormat.OpenXml.Spreadsheet.Font(new Bold(), new FontSize { Val = 10 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFFFFFFF" }, new FontName { Val = "Calibri" }), // 3 th
            new DocumentFormat.OpenXml.Spreadsheet.Font(new Bold(), new FontSize { Val = 10 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF141414" }, new FontName { Val = "Calibri" }), // 4 bold
            new DocumentFormat.OpenXml.Spreadsheet.Font(new Bold(), new FontSize { Val = 9 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF56613F" }, new FontName { Val = "Calibri" }), // 5 kpi label
            new DocumentFormat.OpenXml.Spreadsheet.Font(new Bold(), new FontSize { Val = 14 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF141414" }, new FontName { Val = "Calibri" }), // 6 kpi val
            new DocumentFormat.OpenXml.Spreadsheet.Font(new FontSize { Val = 10 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFB91C1C" }, new FontName { Val = "Calibri" }), // 7 red
            new DocumentFormat.OpenXml.Spreadsheet.Font(new Bold(), new FontSize { Val = 11 }, new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF56613F" }, new FontName { Val = "Calibri" }) // 8 section
        );

        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }), // 0
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), // 1
            new Fill(new PatternFill(new ForegroundColor { Rgb = "FF56613F" }) { PatternType = PatternValues.Solid }), // 2 olive header
            new Fill(new PatternFill(new ForegroundColor { Rgb = "FFF4F4F1" }) { PatternType = PatternValues.Solid }), // 3 soft kpi
            new Fill(new PatternFill(new ForegroundColor { Rgb = "FFEEF2E3" }) { PatternType = PatternValues.Solid })  // 4 soft olive
        );

        var thinBorder = new Border(
            new LeftBorder(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFD9D9D4" }) { Style = BorderStyleValues.Thin },
            new RightBorder(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFD9D9D4" }) { Style = BorderStyleValues.Thin },
            new TopBorder(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFD9D9D4" }) { Style = BorderStyleValues.Thin },
            new BottomBorder(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFD9D9D4" }) { Style = BorderStyleValues.Thin },
            new DiagonalBorder()
        );
        var borders = new Borders(new Border(), thinBorder);

        var numFmts = new NumberingFormats(
            new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat { NumberFormatId = 164, FormatCode = "#,##0 \"đ\"" },
            new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat { NumberFormatId = 165, FormatCode = "0.00%" },
            new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat { NumberFormatId = 166, FormatCode = "#,##0" }
        );

        var cellFormats = new CellFormats(
            new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }, // 0 ST_DEFAULT
            new CellFormat { FontId = 1, FillId = 0, BorderId = 0 }, // 1 ST_TITLE
            new CellFormat { FontId = 2, FillId = 0, BorderId = 0 }, // 2 ST_SUBTITLE
            new CellFormat { FontId = 3, FillId = 2, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, Alignment = new Alignment { Vertical = VerticalAlignmentValues.Center } }, // 3 ST_TH
            new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // 4 ST_TEXT
            new CellFormat { FontId = 4, FillId = 0, BorderId = 1, ApplyFont = true, ApplyBorder = true }, // 5 ST_TEXT_BOLD
            new CellFormat { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 166, ApplyNumberFormat = true, ApplyBorder = true }, // 6 ST_NUM
            new CellFormat { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 164, ApplyNumberFormat = true, ApplyBorder = true }, // 7 ST_CURRENCY
            new CellFormat { FontId = 0, FillId = 0, BorderId = 1, NumberFormatId = 165, ApplyNumberFormat = true, ApplyBorder = true }, // 8 ST_PCT
            new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }, // 9 ST_DATE
            new CellFormat { FontId = 5, FillId = 4, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }, // 10 ST_KPI_LABEL
            new CellFormat { FontId = 6, FillId = 3, BorderId = 1, NumberFormatId = 164, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyNumberFormat = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center } }, // 11 ST_KPI_VAL_CUR
            new CellFormat { FontId = 6, FillId = 3, BorderId = 1, NumberFormatId = 166, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyNumberFormat = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center } }, // 12 ST_KPI_VAL_NUM
            new CellFormat { FontId = 6, FillId = 3, BorderId = 1, NumberFormatId = 165, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyNumberFormat = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center } }, // 13 ST_KPI_VAL_PCT
            new CellFormat { FontId = 7, FillId = 0, BorderId = 1, NumberFormatId = 164, ApplyFont = true, ApplyNumberFormat = true, ApplyBorder = true }, // 14 ST_NEG_CUR
            new CellFormat { FontId = 8, FillId = 0, BorderId = 0, ApplyFont = true } // 15 ST_SECTION
        );

        return new Stylesheet(numFmts, fonts, fills, borders, cellFormats);
    }
}
