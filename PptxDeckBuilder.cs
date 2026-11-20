using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

public static class PptxDeckBuilder
{
    // 16:9 widescreen (13.333 in x 7.5 in)
    private const long SLIDE_W = 12192000L;
    private const long SLIDE_H = 6858000L;
    private const long EMU_PER_INCH = 914400L;

    // Palette matching Ngọc Linh Data Studio web UI
    private const string C_BG = "F4F4F1";
    private const string C_CARD = "FFFFFF";
    private const string C_SOFT = "EAEAE5";
    private const string C_INK = "141414";
    private const string C_MUTED = "6E6E69";
    private const string C_OLIVE = "56613F";
    private const string C_OLIVE_DARK = "3C4430";
    private const string C_OLIVE_SOFT = "EEF2E3";
    private const string C_AMBER = "B7791F";
    private const string C_AMBER_SOFT = "FEF3C7";
    private const string C_BORDER = "DCDCD6";
    private const string FONT = "Segoe UI";

    private static long In(double inches) => (long)(inches * EMU_PER_INCH);

    public static void Build(string outputPath, string screenshotsDir)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        using var doc = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);
        var presPart = doc.AddPresentationPart();
        presPart.Presentation = new Presentation();

        var slideMasterPart = presPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup())
            )),
            new ColorMapOverride(new A.MasterColorMapping())
        );
        slideLayoutPart.SlideLayout.Save();

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup())
            )),
            new P.ColorMap { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) })
        );
        slideMasterPart.SlideMaster.Save();

        var themePart = slideMasterPart.AddNewPart<ThemePart>();
        themePart.Theme = BuildTheme();
        themePart.Theme.Save();

        presPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
            new SlideMasterId { Id = 2147483648U, RelationshipId = presPart.GetIdOfPart(slideMasterPart) }
        );
        var slideIdList = presPart.Presentation.AppendChild(new SlideIdList());
        presPart.Presentation.SlideSize = new SlideSize { Cx = (int)SLIDE_W, Cy = (int)SLIDE_H, Type = SlideSizeValues.Screen16x9 };
        presPart.Presentation.NotesSize = new NotesSize { Cx = 6858000, Cy = 9144000 };

        uint slideId = 256;
        BuildSlide1(presPart, slideLayoutPart, slideIdList, ref slideId, Path.Combine(screenshotsDir, "home.png"));
        BuildSlide2(presPart, slideLayoutPart, slideIdList, ref slideId);
        BuildSlide3(presPart, slideLayoutPart, slideIdList, ref slideId);
        BuildSlide4(presPart, slideLayoutPart, slideIdList, ref slideId, screenshotsDir);
        BuildSlide5(presPart, slideLayoutPart, slideIdList, ref slideId);

        presPart.Presentation.Save();
    }

    // ================= SLIDE 1: GIỚI THIỆU TỔNG QUAN =================
    private static void BuildSlide1(PresentationPart presPart, SlideLayoutPart layoutPart, SlideIdList idList, ref uint slideId, string homeImg)
    {
        var (sp, tree) = NewSlide(presPart, layoutPart, idList, ref slideId);
        uint id = 10;

        // Top tag pill
        AddCard(tree, ref id, 0.55, 0.42, 4.3, 0.38, C_OLIVE_SOFT, C_OLIVE_SOFT, 18000);
        AddText(tree, ref id, 0.68, 0.46, 4.05, 0.30, "BI DATA CLEANSING & ANALYTICS WORKBENCH", 10, true, C_OLIVE);

        // Title & Subtitle
        AddText(tree, ref id, 0.55, 0.92, 6.2, 0.65, "NGỌC LINH DATA STUDIO", 30, true, C_INK);
        AddText(tree, ref id, 0.55, 1.52, 6.2, 0.45, "Nền tảng làm sạch, phân tích dữ liệu và hỗ trợ ra quyết định kinh doanh.", 13, false, C_MUTED);

        // Info Card 1: Hệ thống & Vấn đề giải quyết
        AddCard(tree, ref id, 0.55, 2.12, 6.0, 1.95, C_CARD, C_BORDER);
        AddText(tree, ref id, 0.78, 2.26, 5.5, 0.32, "Hệ thống giải quyết vấn đề gì?", 13, true, C_OLIVE_DARK);
        AddBullets(tree, ref id, 0.78, 2.62, 5.5, 1.35, 11.5, C_INK, new[]
        {
            "Loại bỏ sai lệch do trùng lặp đơn hàng, đơn hủy và thiếu chuẩn hóa.",
            "Thay thế thao tác ghép bảng, đối soát thủ công tốn thời gian.",
            "Chuyển đổi dữ liệu bán hàng thô (Excel/CSV) thành KPI & báo cáo quản trị."
        });

        // Info Card 2: Đối tượng & Giá trị cốt lõi
        AddCard(tree, ref id, 0.55, 4.22, 6.0, 1.68, C_CARD, C_BORDER);
        AddText(tree, ref id, 0.78, 4.36, 5.5, 0.32, "Đối tượng sử dụng & Giá trị cốt lõi", 13, true, C_OLIVE_DARK);
        AddBullets(tree, ref id, 0.78, 4.72, 5.5, 1.05, 11.5, C_INK, new[]
        {
            "Đối tượng: Chuyên viên BI, Data Analyst, Business Analyst, Nhà quản trị.",
            "Giá trị: Tự động hóa quy trình BI · Nâng cao chất lượng dữ liệu · Ra quyết định nhanh."
        });

        // Key message olive banner at bottom left
        AddCard(tree, ref id, 0.55, 6.08, 6.0, 0.92, C_OLIVE, C_OLIVE);
        AddText(tree, ref id, 0.78, 6.22, 5.55, 0.65, "“Biến dữ liệu thô thành thông tin có giá trị cho quyết định kinh doanh.”", 13.5, true, "FFFFFF");

        // Right column: Real screenshot frame
        AddCard(tree, ref id, 6.78, 0.55, 6.0, 6.45, C_CARD, C_BORDER);
        AddText(tree, ref id, 7.0, 0.72, 5.55, 0.32, "Giao diện thực tế trên Localhost (http://localhost:5000)", 11.5, true, C_OLIVE);
        if (File.Exists(homeImg))
            AddImage(sp, tree, ref id, homeImg, 6.95, 1.15, 5.66, 3.54);
        AddCard(tree, ref id, 6.95, 4.86, 5.66, 1.92, C_SOFT, C_SOFT);
        AddText(tree, ref id, 7.15, 5.0, 5.25, 0.3, "Điểm nổi bật của phiên bản hiện tại", 12, true, C_INK);
        AddBullets(tree, ref id, 7.15, 5.34, 5.25, 1.35, 11, C_MUTED, new[]
        {
            "Vận hành sạch từ dữ liệu thực tế của khách hàng (không tạo số liệu giả).",
            "Quy trình 4 bước liền mạch kết hợp Cổng con người kiểm duyệt.",
            "Xuất trực tiếp Excel BI Workbook (.xlsx) & Báo cáo Word BI (.docx)."
        });

        sp.Slide.Save();
    }

    // ================= SLIDE 2: NĂNG LỰC CỦA HỆ THỐNG =================
    private static void BuildSlide2(PresentationPart presPart, SlideLayoutPart layoutPart, SlideIdList idList, ref uint slideId)
    {
        var (sp, tree) = NewSlide(presPart, layoutPart, idList, ref slideId);
        uint id = 10;

        AddSlideHeader(tree, ref id, "SLIDE 2 · NĂNG LỰC HỆ THỐNG", "NĂNG LỰC CỦA NGỌC LINH DATA STUDIO", "Tổng hợp 4 nhóm chức năng cốt lõi đã triển khai và lộ trình mở rộng");

        var cards = new[]
        {
            new {
                X = 0.55, Y = 1.45, Badge = "1. DATA INTEGRATION", Title = "Kết nối dữ liệu",
                Status = "Đang hoạt động", StatusBg = C_OLIVE_SOFT, StatusFg = C_OLIVE,
                Bullets = new[] {
                    "Tiếp nhận tệp Excel (.xlsx, .xls), CSV và JSON nhiều tệp cùng lúc.",
                    "Tự động kiểm tra cấu trúc cột (tiếng Việt & tiếng Anh) và định dạng.",
                    "Quản lý danh sách tệp: Thêm, Thay thế, Xóa và tải Tệp Template."
                }
            },
            new {
                X = 6.75, Y = 1.45, Badge = "2. DATA CLEANSING", Title = "Làm sạch dữ liệu",
                Status = "Đang hoạt động", StatusBg = C_OLIVE_SOFT, StatusFg = C_OLIVE,
                Bullets = new[] {
                    "Phát hiện trùng lặp theo khóa (Mã đơn + SKU + Sản phẩm).",
                    "Loại bỏ đơn hủy/lỗi (CANCELLED, FAILED) & giảm trừ hàng hoàn.",
                    "Chuẩn hóa ngày/số/chữ; chỉnh sửa trực tiếp bản ghi ngoại lệ & Phê duyệt."
                }
            },
            new {
                X = 0.55, Y = 3.72, Badge = "3. DATA ANALYTICS", Title = "Phân tích dữ liệu",
                Status = "Đang hoạt động", StatusBg = C_OLIVE_SOFT, StatusFg = C_OLIVE,
                Bullets = new[] {
                    "Tính toán KPI: Doanh thu thuần, Lợi nhuận gộp, Biên LN %, AOV.",
                    "Phân tích theo Sản phẩm, Chi nhánh, Kênh, Chiết khấu & Thời gian.",
                    "Sinh Insight 5 nhóm & Trợ lý Q&A dựa trên bộ quy tắc lập trình."
                }
            },
            new {
                X = 6.75, Y = 3.72, Badge = "4. BI REPORTING", Title = "Báo cáo BI đầu ra",
                Status = "Đang hoạt động", StatusBg = C_OLIVE_SOFT, StatusFg = C_OLIVE,
                Bullets = new[] {
                    "Xuất Excel BI Workbook (.xlsx): 6 Sheets, công thức động & 4 biểu đồ.",
                    "Xuất Word BI Report (.docx): 11 phần, bảng biểu & 3 biểu đồ PNG.",
                    "Hỗ trợ xem trước cấu trúc báo cáo, xuất CSV, In PDF và JSON."
                }
            }
        };

        foreach (var c in cards)
        {
            AddCard(tree, ref id, c.X, c.Y, 6.02, 2.08, C_CARD, C_BORDER);
            AddText(tree, ref id, c.X + 0.24, c.Y + 0.16, 3.4, 0.25, c.Badge, 10, true, C_OLIVE);
            AddCard(tree, ref id, c.X + 4.35, c.Y + 0.14, 1.45, 0.28, c.StatusBg, c.StatusBg, 14000);
            AddText(tree, ref id, c.X + 4.42, c.Y + 0.17, 1.32, 0.24, c.Status, 9.5, true, c.StatusFg, A.TextAlignmentTypeValues.Center);
            AddText(tree, ref id, c.X + 0.24, c.Y + 0.44, 5.5, 0.34, c.Title, 15, true, C_INK);
            AddBullets(tree, ref id, c.X + 0.24, c.Y + 0.82, 5.5, 1.15, 11, C_MUTED, c.Bullets);
        }

        // Bottom roadmap distinction bar (Active vs Developing vs Planned)
        AddCard(tree, ref id, 0.55, 6.0, 12.22, 1.02, C_INK, C_INK);
        AddText(tree, ref id, 0.78, 6.1, 3.8, 0.26, "● ĐANG HOẠT ĐỘNG", 10.5, true, "C9D79E");
        AddText(tree, ref id, 0.78, 6.38, 3.7, 0.52, "Toàn bộ quy trình 4 bước BI, làm sạch, phân tích thống kê & xuất .xlsx/.docx thực tế.", 10, false, "E5E5E0");

        AddText(tree, ref id, 4.78, 6.1, 3.8, 0.26, "◐ ĐANG PHÁT TRIỂN", 10.5, true, "F5D46A");
        AddText(tree, ref id, 4.78, 6.38, 3.7, 0.52, "Mở rộng nhận diện đa mẫu biểu doanh nghiệp & tinh chỉnh bộ quy tắc ngoại lệ.", 10, false, "E5E5E0");

        AddText(tree, ref id, 8.78, 6.1, 3.8, 0.26, "○ DỰ KIẾN PHÁT TRIỂN", 10.5, true, "D4D4CE");
        AddText(tree, ref id, 8.78, 6.38, 3.7, 0.52, "Kết nối trực tiếp Google Sheets / ERP API & tích hợp LLM (Gemini/OpenAI) phân tích tự do.", 10, false, "E5E5E0");

        sp.Slide.Save();
    }

    // ================= SLIDE 3: KIẾN TRÚC VÀ CÁCH THỨC HOẠT ĐỘNG =================
    private static void BuildSlide3(PresentationPart presPart, SlideLayoutPart layoutPart, SlideIdList idList, ref uint slideId)
    {
        var (sp, tree) = NewSlide(presPart, layoutPart, idList, ref slideId);
        uint id = 10;

        AddSlideHeader(tree, ref id, "SLIDE 3 · KIẾN TRÚC HỆ THỐNG", "KIẾN TRÚC VÀ CÁCH THỨC HOẠT ĐỘNG CỦA AGENT", "Mô hình 6 lớp từ nguồn dữ liệu đầu vào đến bộ báo cáo BI (đối chiếu 100% source code)");

        // Layer 1: Input
        AddArchBox(tree, ref id, 0.55, 1.55, 2.2, 3.6, C_CARD, C_BORDER, "1. INPUT LAYER", "Nguồn dữ liệu đầu vào", new[]
        {
            "Excel (.xlsx, .xls) đọc qua SheetJS trên trình duyệt.",
            "CSV & JSON gửi qua /api/connect/upload-file.",
            "Tệp Template CSV chuẩn hóa 12 cột."
        });
        AddArrow(tree, ref id, 2.82, 3.2, 0.35, 0.3);

        // Layer 2: Processing
        AddArchBox(tree, ref id, 3.22, 1.55, 2.65, 3.6, C_CARD, C_BORDER, "2. DATA PROCESSING", "Xử lý & Làm sạch", new[]
        {
            "C# DataAnalystHub: nhận diện cột, chuẩn hóa ngày/số/trạng thái.",
            "Khử trùng lặp, loại đơn hủy, đối soát hoàn trả, nội suy giá vốn 75%.",
            "Xuất bảng ngoại lệ cho phép sửa trực tiếp."
        });
        AddArrow(tree, ref id, 5.94, 3.2, 0.35, 0.3);

        // Layer 3: Analysis
        AddArchBox(tree, ref id, 6.34, 1.55, 2.65, 3.6, C_CARD, C_BORDER, "3. ANALYSIS LAYER", "Phân tích & KPI", new[]
        {
            "PerformDeepAnalysis(): tính Doanh thu thuần, Giá vốn, LN gộp, Biên LN.",
            "Phân tổ Chi nhánh, Sản phẩm, Kênh, Chiết khấu & Chuỗi thời gian.",
            "Tổng hợp Insight theo 5 nhóm nghiệp vụ."
        });
        AddArrow(tree, ref id, 9.06, 3.2, 0.35, 0.3);

        // Layer 5: Output
        AddArchBox(tree, ref id, 9.46, 1.55, 3.3, 3.6, C_OLIVE, C_OLIVE, "5. OUTPUT LAYER", "Sản phẩm đầu ra thực tế", new[]
        {
            "ExcelBiExporter.cs (OpenXML): tạo .xlsx 6 Sheets, công thức động & 4 biểu đồ.",
            "WordBiExporter.cs (OpenXML + System.Drawing): tạo .docx A4 11 phần & 3 biểu đồ PNG.",
            "Hỗ trợ xuất thêm CSV, In PDF & JSON."
        }, true);

        // Layer 4: AI / Agent Layer (bottom left-center)
        AddCard(tree, ref id, 0.55, 5.4, 6.2, 1.65, C_CARD, C_OLIVE);
        AddText(tree, ref id, 0.78, 5.52, 5.8, 0.28, "4. AI / AGENT LAYER (RULE-BASED & HUMAN-IN-THE-LOOP)", 10.5, true, C_OLIVE);
        AddBullets(tree, ref id, 0.78, 5.84, 5.8, 1.1, 10.5, C_INK, new[]
        {
            "Vận hành bằng động cơ quy tắc lập trình & thống kê (chưa kết nối LLM tự chủ).",
            "Q&A Copilot (/api/analyze/ask-copilot) trả lời theo từ khóa trên số liệu sạch.",
            "2 Cổng con người duyệt: Kiểm duyệt dữ liệu sạch & Đánh giá kết quả phân tích."
        });

        // Layer 6: User Interface (bottom right)
        AddCard(tree, ref id, 6.95, 5.4, 5.81, 1.65, C_SOFT, C_BORDER);
        AddText(tree, ref id, 7.18, 5.52, 5.4, 0.28, "6. USER INTERFACE & BACKEND HOSTING", 10.5, true, C_OLIVE_DARK);
        AddBullets(tree, ref id, 7.18, 5.84, 5.4, 1.1, 10.5, C_INK, new[]
        {
            "Frontend: Single-Page App (HTML5, Tailwind CSS, Phosphor Icons, SheetJS).",
            "Backend: ASP.NET Core 8.0 Minimal API chạy trên Kestrel (localhost:5000).",
            "Thư viện sinh tệp: DocumentFormat.OpenXml 3.1.1 & System.Drawing.Common 8.0.10."
        });

        sp.Slide.Save();
    }

    // ================= SLIDE 4: QUY TRÌNH 4 BƯỚC BI =================
    private static void BuildSlide4(PresentationPart presPart, SlideLayoutPart layoutPart, SlideIdList idList, ref uint slideId, string shotsDir)
    {
        var (sp, tree) = NewSlide(presPart, layoutPart, idList, ref slideId);
        uint id = 10;

        AddSlideHeader(tree, ref id, "SLIDE 4 · HƯỚNG DẪN SỬ DỤNG", "HƯỚNG DẪN SỬ DỤNG – 4 BƯỚC BI", "Thao tác của người dùng và xử lý tự động của hệ thống qua từng bước");

        var steps = new[]
        {
            new {
                Num = "BƯỚC 1", Title = "Kết nối dữ liệu", Img = "step1.png",
                User = "Truy cập localhost, tải lên tệp Excel/CSV; thêm, thay thế hoặc xóa tệp.",
                System = "Đọc tệp, kiểm tra cấu trúc cột, đếm bản ghi & cảnh báo sơ bộ."
            },
            new {
                Num = "BƯỚC 2", Title = "Làm sạch dữ liệu", Img = "step2.png",
                User = "Cấu hình quy tắc, sửa bản ghi ngoại lệ trực tiếp & nhấn Phê duyệt.",
                System = "Khử trùng, loại đơn hủy, giảm trừ hàng hoàn, chuẩn hóa & ghi log."
            },
            new {
                Num = "BƯỚC 3", Title = "Phân tích dữ liệu", Img = "step3.png",
                User = "Theo dõi Dashboard, hỏi đáp với AI & nhấn Đánh giá kết quả.",
                System = "Tính KPI, vẽ xu hướng thời gian, phân tích chi nhánh/SP & sinh Insight."
            },
            new {
                Num = "BƯỚC 4", Title = "Xuất kết quả", Img = "step4.png",
                User = "Nhấn Xem trước & tải Excel BI (.xlsx) hoặc Báo cáo Word (.docx).",
                System = "Đồng bộ số liệu sạch để sinh Workbook 6 Sheets & Báo cáo 11 phần."
            }
        };

        for (int i = 0; i < 4; i++)
        {
            double x = 0.55 + i * 3.12;
            var s = steps[i];
            AddCard(tree, ref id, x, 1.42, 2.92, 5.6, C_CARD, C_BORDER);

            // Step pill
            AddCard(tree, ref id, x + 0.18, 1.56, 1.0, 0.28, C_OLIVE, C_OLIVE, 14000);
            AddText(tree, ref id, x + 0.18, 1.58, 1.0, 0.24, s.Num, 10, true, "FFFFFF", A.TextAlignmentTypeValues.Center);
            AddText(tree, ref id, x + 0.18, 1.9, 2.56, 0.34, s.Title, 14, true, C_INK);

            string imgPath = Path.Combine(shotsDir, s.Img);
            if (File.Exists(imgPath))
                AddImage(sp, tree, ref id, imgPath, x + 0.18, 2.32, 2.56, 1.6);

            AddCard(tree, ref id, x + 0.18, 4.06, 2.56, 1.32, C_SOFT, C_SOFT);
            AddText(tree, ref id, x + 0.28, 4.14, 2.36, 0.24, "Người dùng thao tác:", 10.5, true, C_OLIVE_DARK);
            AddText(tree, ref id, x + 0.28, 4.4, 2.36, 0.9, s.User, 10.5, false, C_INK);

            AddCard(tree, ref id, x + 0.18, 5.48, 2.56, 1.38, C_OLIVE_SOFT, C_OLIVE_SOFT);
            AddText(tree, ref id, x + 0.28, 5.56, 2.36, 0.24, "Hệ thống thực hiện:", 10.5, true, C_OLIVE);
            AddText(tree, ref id, x + 0.28, 5.82, 2.36, 0.95, s.System, 10.5, false, C_INK);

            if (i < 3) AddArrow(tree, ref id, x + 2.94, 3.0, 0.16, 0.24);
        }

        sp.Slide.Save();
    }

    // ================= SLIDE 5: SẢN PHẨM ĐẦU RA & GIÁ TRỊ =================
    private static void BuildSlide5(PresentationPart presPart, SlideLayoutPart layoutPart, SlideIdList idList, ref uint slideId)
    {
        var (sp, tree) = NewSlide(presPart, layoutPart, idList, ref slideId);
        uint id = 10;

        AddSlideHeader(tree, ref id, "SLIDE 5 · SẢN PHẨM ĐẦU RA & GIÁ TRỊ", "TỪ DỮ LIỆU ĐẾN QUYẾT ĐỊNH KINH DOANH", "Hai sản phẩm đầu ra chuẩn doanh nghiệp được tạo tự động từ dữ liệu thực tế");

        // Product 1: Excel
        AddCard(tree, ref id, 0.55, 1.42, 6.0, 3.55, C_CARD, C_BORDER);
        AddCard(tree, ref id, 0.78, 1.6, 2.4, 0.28, C_OLIVE_SOFT, C_OLIVE_SOFT, 14000);
        AddText(tree, ref id, 0.85, 1.63, 2.26, 0.24, "SẢN PHẨM 1 · .XLSX", 10, true, C_OLIVE);
        AddText(tree, ref id, 0.78, 1.95, 5.5, 0.36, "EXCEL BI WORKBOOK (6 SHEETS)", 15, true, C_INK);
        AddBullets(tree, ref id, 0.78, 2.38, 5.5, 2.45, 11.5, C_INK, new[]
        {
            "RAW DATA: Lưu nguyên trạng dữ liệu gốc ban đầu.",
            "CLEANED DATA: Excel Table sạch, có bộ lọc & công thức tính KPI.",
            "DATA DICTIONARY: Từ điển 18 trường dữ liệu chuẩn BI.",
            "CLEANING LOG: Thống kê trước/sau & chi tiết ngoại lệ.",
            "ANALYSIS: Bảng tổng hợp động (SUMIFS, COUNTIFS) theo nhóm.",
            "DASHBOARD: 5 KPI Cards, Insight & 4 biểu đồ Excel thực tế."
        });

        // Product 2: Word
        AddCard(tree, ref id, 6.78, 1.42, 6.0, 3.55, C_CARD, C_BORDER);
        AddCard(tree, ref id, 7.01, 1.6, 2.4, 0.28, C_INK, C_INK, 14000);
        AddText(tree, ref id, 7.08, 1.63, 2.26, 0.24, "SẢN PHẨM 2 · .DOCX", 10, true, "FFFFFF");
        AddText(tree, ref id, 7.01, 1.95, 5.5, 0.36, "BI ANALYSIS REPORT (11 PHẦN)", 15, true, C_INK);
        AddBullets(tree, ref id, 7.01, 2.38, 5.5, 2.45, 11.5, C_INK, new[]
        {
            "Trang bìa, Mục lục tự động, Executive Summary dành cho Ban Giám đốc.",
            "Tổng quan, đánh giá chất lượng dữ liệu & Phương pháp tính KPI.",
            "Kết quả phân tích kinh doanh kèm bảng số liệu & 3 biểu đồ PNG.",
            "Business Insights (Hiện tượng → Bằng chứng → Phân tích → Tác động).",
            "Đề xuất giải pháp theo 3 mức ưu tiên & Kế hoạch triển khai.",
            "Kết luận hỗ trợ ra quyết định & Phụ lục từ điển dữ liệu."
        });

        // 4 Value Cards
        var vals = new[]
        {
            ("1. Tiết kiệm thời gian", "Tự động hóa đối soát & lập báo cáo định kỳ."),
            ("2. Hạn chế sai sót", "Loại bỏ đơn trùng, đơn hủy & kiểm soát ngoại lệ."),
            ("3. Nâng cao phân tích", "Khai thác đa chiều theo SP, chi nhánh, thời gian."),
            ("4. Hỗ trợ quyết định", "Cung cấp KPI & giải pháp dựa trên dữ liệu thực.")
        };
        for (int i = 0; i < 4; i++)
        {
            double x = 0.55 + i * 3.12;
            AddCard(tree, ref id, x, 5.14, 2.92, 1.0, C_SOFT, C_SOFT);
            AddText(tree, ref id, x + 0.16, 5.24, 2.6, 0.28, vals[i].Item1, 11.5, true, C_OLIVE_DARK);
            AddText(tree, ref id, x + 0.16, 5.54, 2.6, 0.52, vals[i].Item2, 10.5, false, C_MUTED);
        }

        // Closing message banner
        AddCard(tree, ref id, 0.55, 6.28, 12.22, 0.72, C_OLIVE, C_OLIVE);
        AddText(tree, ref id, 0.78, 6.44, 11.76, 0.4, "“Ngọc Linh Data Studio – Từ dữ liệu thô đến giá trị kinh doanh.”", 15, true, "FFFFFF", A.TextAlignmentTypeValues.Center);

        sp.Slide.Save();
    }

    // ================= PRIMITIVES =================
    private static (SlidePart, ShapeTree) NewSlide(PresentationPart presPart, SlideLayoutPart layoutPart, SlideIdList idList, ref uint slideId)
    {
        var sp = presPart.AddNewPart<SlidePart>();
        sp.AddPart(layoutPart);

        var tree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup())
        );

        sp.Slide = new Slide(
            new CommonSlideData(
                new Background(new BackgroundProperties(new A.SolidFill(new A.RgbColorModelHex { Val = C_BG }), new A.EffectList())),
                tree
            ),
            new ColorMapOverride(new A.MasterColorMapping())
        );

        idList.Append(new SlideId { Id = slideId++, RelationshipId = presPart.GetIdOfPart(sp) });
        return (sp, tree);
    }

    private static void AddSlideHeader(ShapeTree tree, ref uint id, string eyebrow, string title, string sub)
    {
        AddText(tree, ref id, 0.55, 0.28, 6.0, 0.24, eyebrow, 10, true, C_OLIVE);
        AddText(tree, ref id, 0.55, 0.52, 11.5, 0.46, title, 23, true, C_INK);
        AddText(tree, ref id, 0.55, 0.98, 11.5, 0.3, sub, 11.5, false, C_MUTED);
    }

    private static void AddArchBox(ShapeTree tree, ref uint id, double x, double y, double w, double h, string bg, string border, string tag, string title, string[] bullets, bool dark = false)
    {
        AddCard(tree, ref id, x, y, w, h, bg, border);
        AddText(tree, ref id, x + 0.16, y + 0.16, w - 0.32, 0.24, tag, 9.5, true, dark ? "C9D79E" : C_OLIVE);
        AddText(tree, ref id, x + 0.16, y + 0.42, w - 0.32, 0.38, title, 13, true, dark ? "FFFFFF" : C_INK);
        AddBullets(tree, ref id, x + 0.16, y + 0.88, w - 0.32, h - 1.0, 10.5, dark ? "F4F4F1" : C_MUTED, bullets);
    }

    private static void AddCard(ShapeTree tree, ref uint id, double x, double y, double w, double h, string fillHex, string borderHex, int radiusAdjust = 6000)
    {
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id++, Name = $"Card{id}" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = In(x), Y = In(y) }, new A.Extents { Cx = In(w), Cy = In(h) }),
                new A.PresetGeometry(new A.AdjustValueList(new A.ShapeGuide { Name = "adj", Formula = $"val {radiusAdjust}" })) { Preset = A.ShapeTypeValues.RoundRectangle },
                new A.SolidFill(new A.RgbColorModelHex { Val = fillHex }),
                new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = borderHex })) { Width = 9525 }
            )
        );
        tree.Append(shape);
    }

    private static void AddArrow(ShapeTree tree, ref uint id, double x, double y, double w, double h)
    {
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id++, Name = $"Arrow{id}" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = In(x), Y = In(y) }, new A.Extents { Cx = In(w), Cy = In(h) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.RightArrow },
                new A.SolidFill(new A.RgbColorModelHex { Val = C_OLIVE }),
                new A.Outline(new A.NoFill())
            )
        );
        tree.Append(shape);
    }

    private static void AddText(ShapeTree tree, ref uint id, double x, double y, double w, double h, string text, double pt, bool bold, string hexColor, A.TextAlignmentTypeValues? align = null)
    {
        var pPr = new A.ParagraphProperties { Alignment = align ?? A.TextAlignmentTypeValues.Left };
        var rPr = new A.RunProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = hexColor }),
            new A.LatinFont { Typeface = FONT },
            new A.ComplexScriptFont { Typeface = FONT }
        ) { Language = "vi-VN", FontSize = (int)(pt * 100), Bold = bold };

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id++, Name = $"Txt{id}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = In(x), Y = In(y) }, new A.Extents { Cx = In(w), Cy = In(h) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()
            ),
            new P.TextBody(
                new A.BodyProperties { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0 },
                new A.ListStyle(),
                new A.Paragraph(pPr, new A.Run(rPr, new A.Text(text)))
            )
        );
        tree.Append(shape);
    }

    private static void AddBullets(ShapeTree tree, ref uint id, double x, double y, double w, double h, double pt, string hexColor, string[] lines)
    {
        var tb = new P.TextBody(
            new A.BodyProperties { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0 },
            new A.ListStyle()
        );
        foreach (var line in lines)
        {
            var rPr = new A.RunProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = hexColor }),
                new A.LatinFont { Typeface = FONT },
                new A.ComplexScriptFont { Typeface = FONT }
            ) { Language = "vi-VN", FontSize = (int)(pt * 100) };

            tb.Append(new A.Paragraph(
                new A.ParagraphProperties(new A.SpaceAfter(new A.SpacingPoints { Val = 450 })),
                new A.Run(rPr, new A.Text("• " + line))
            ));
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id++, Name = $"Bul{id}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = In(x), Y = In(y) }, new A.Extents { Cx = In(w), Cy = In(h) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()
            ),
            tb
        );
        tree.Append(shape);
    }

    private static void AddImage(SlidePart sp, ShapeTree tree, ref uint id, string imgPath, double x, double y, double w, double h)
    {
        var imgPart = sp.AddImagePart(ImagePartType.Png);
        using (var fs = File.OpenRead(imgPath)) imgPart.FeedData(fs);
        string relId = sp.GetIdOfPart(imgPart);

        var pic = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id++, Name = Path.GetFileName(imgPath) },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()
            ),
            new P.BlipFill(
                new A.Blip { Embed = relId },
                new A.Stretch(new A.FillRectangle())
            ),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = In(x), Y = In(y) }, new A.Extents { Cx = In(w), Cy = In(h) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = C_BORDER })) { Width = 9525 }
            )
        );
        tree.Append(pic);
    }

    private static A.Theme BuildTheme() =>
        new(new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = C_OLIVE_DARK }),
                new A.Light2Color(new A.RgbColorModelHex { Val = C_BG }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = C_OLIVE }),
                new A.Accent2Color(new A.RgbColorModelHex { Val = C_AMBER }),
                new A.Accent3Color(new A.RgbColorModelHex { Val = "7A8760" }),
                new A.Accent4Color(new A.RgbColorModelHex { Val = "A7B38C" }),
                new A.Accent5Color(new A.RgbColorModelHex { Val = C_INK }),
                new A.Accent6Color(new A.RgbColorModelHex { Val = C_MUTED }),
                new A.Hyperlink(new A.RgbColorModelHex { Val = C_OLIVE }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = C_MUTED })
            ) { Name = "NgocLinhTheme" },
            new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = FONT }, new A.EastAsianFont { Typeface = "" }, new A.ComplexScriptFont { Typeface = FONT }),
                new A.MinorFont(new A.LatinFont { Typeface = FONT }, new A.EastAsianFont { Typeface = "" }, new A.ComplexScriptFont { Typeface = FONT })
            ) { Name = "NgocLinhFonts" },
            new A.FormatScheme(
                new A.FillStyleList(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }), new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }), new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                new A.LineStyleList(new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 }, new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 }, new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 }),
                new A.EffectStyleList(new A.EffectStyle(new A.EffectList()), new A.EffectStyle(new A.EffectList()), new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }), new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }), new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }))
            ) { Name = "NgocLinhFmt" }
        )) { Name = "NgocLinhTheme" };
}
