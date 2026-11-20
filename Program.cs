using System.Globalization;
using System.Text;
using System.Text.Json;

if (args.Contains("--build-slides"))
{
    var outPptx = Path.Combine(Directory.GetCurrentDirectory(), "output", "Ngoc_Linh_Data_Studio_Product_Overview.pptx");
    var shotsDir = Path.Combine(Path.GetTempPath(), "slides_shots");
    PptxDeckBuilder.Build(outPptx, shotsDir);
    Console.WriteLine($"BUILT_PPTX={outPptx}");
    return;
}

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseUrls("http://localhost:5000");
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = 250L * 1024 * 1024;
});
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(o =>
{
    o.MultipartBodyLengthLimit = 250L * 1024 * 1024;
    o.ValueLengthLimit = int.MaxValue;
});
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy => policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod());
});

var app = builder.Build();
app.UseCors();
app.UseDefaultFiles();
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers.CacheControl = "no-cache, no-store, must-revalidate";
        ctx.Context.Response.Headers.Pragma = "no-cache";
        ctx.Context.Response.Headers.Expires = "0";
    }
});

// Khởi động với trạng thái trống - chờ khách hàng tải dữ liệu lên
var dataHub = new DataAnalystHub();

// ================= API ENDPOINTS CHO HỆ THỐNG BI 4 BƯỚC =================

// Overview trạng thái toàn hệ thống
app.MapGet("/api/status", () => Results.Ok(dataHub.GetPipelineStatus()));

app.MapGet("/api/job/progress", () => Results.Ok(dataHub.Progress));

app.MapPost("/api/connect/business-model", (BusinessModelPayload payload) =>
{
    dataHub.BusinessModel = string.IsNullOrWhiteSpace(payload.BusinessModel) ? "Bán lẻ đa kênh" : payload.BusinessModel.Trim();
    return Results.Ok(new { businessModel = dataHub.BusinessModel });
});

// Xoá toàn bộ dữ liệu phiên làm việc
app.MapPost("/api/reset", () =>
{
    dataHub.Reset();
    return Results.Ok(new { message = "Đã xoá dữ liệu. Hệ thống sẵn sàng nhận tệp mới." });
});

// ---------------- BƯỚC 1: KẾT NỐI DỮ LIỆU ----------------
app.MapGet("/api/connect/template", () =>
{
    var header = "OrderCode,OrderDate,CustomerName,StoreName,Sku,ProductName,Quantity,UnitPrice,DiscountAmount,UnitCost,Status,Channel\n";
    var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(header)).ToArray();
    return Results.File(bytes, "text/csv", "DataLens_Template.csv");
});

app.MapGet("/api/connect/files", () => Results.Ok(new
{
    files = dataHub.GetFileSummaries(),
    summary = dataHub.GetRawSummary()
}));

app.MapPost("/api/connect/upload-json", (UploadParsedRequest req) =>
{
    if (req.Records == null || req.Records.Count == 0)
        return Results.BadRequest(new { message = "Dữ liệu tải lên trống." });

    var info = dataHub.AddOrReplaceFile(req.FileName ?? "Uploaded_Data.xlsx", req.Records, req.Headers ?? new List<string>(), req.ReplaceFileId);
    return Results.Ok(new
    {
        message = $"Đã kết nối thành công {req.Records.Count} bản ghi từ '{info.FileName}'!",
        file = info.ToSummary(),
        files = dataHub.GetFileSummaries(),
        summary = dataHub.GetRawSummary(),
        preview = dataHub.RawData.Take(50).ToList()
    });
});

app.MapPost("/api/connect/upload-batch", (UploadBatchRequest req) =>
{
    if (req.Records == null)
        return Results.BadRequest(new { message = "Batch dữ liệu không hợp lệ." });

    var result = dataHub.AppendBatch(req);
    return Results.Ok(result);
});

app.MapPost("/api/connect/upload-file", async (HttpRequest request) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { message = "Yêu cầu định dạng tệp hợp lệ (form-data)." });

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    var replaceId = form["replaceFileId"].FirstOrDefault();
    if (file == null || file.Length == 0)
        return Results.BadRequest(new { message = "Không tìm thấy tệp đính kèm." });

    if (file.Length > 200L * 1024 * 1024)
        return Results.BadRequest(new { message = "Tệp vượt quá giới hạn tối đa 200MB của hệ thống." });

    await using var stream = file.OpenReadStream();
    var (success, info, error) = dataHub.ParseAndAddFileStream(file.FileName, stream, file.Length, replaceId);

    if (!success || info == null)
        return Results.BadRequest(new { message = $"Lỗi đọc tệp: {error}" });

    return Results.Ok(new
    {
        message = $"Đã phân tích và kết nối thành công {info.RecordCount:N0} bản ghi từ tệp '{file.FileName}'!",
        file = info.ToSummary(),
        files = dataHub.GetFileSummaries(),
        summary = dataHub.GetRawSummary(),
        preview = dataHub.RawData.Take(50).ToList(),
        progress = dataHub.Progress
    });
});

app.MapDelete("/api/connect/files/{id}", (string id) =>
{
    var removed = dataHub.RemoveFile(id);
    if (!removed) return Results.NotFound(new { message = "Không tìm thấy tệp cần xóa." });
    return Results.Ok(new
    {
        message = "Đã xóa tệp dữ liệu.",
        files = dataHub.GetFileSummaries(),
        summary = dataHub.GetRawSummary(),
        preview = dataHub.RawData.Take(50).ToList()
    });
});

app.MapGet("/api/connect/preview", () => Results.Ok(new
{
    files = dataHub.GetFileSummaries(),
    summary = dataHub.GetRawSummary(),
    records = dataHub.RawData.Take(50).ToList()
}));

app.MapGet("/api/connect/records", (int page = 1, int pageSize = 50, string? search = null) =>
{
    page = Math.Max(1, page);
    pageSize = Math.Clamp(pageSize, 10, 200);
    IEnumerable<GenericDataRecord> q = dataHub.RawData;
    if (!string.IsNullOrWhiteSpace(search))
    {
        var s = search.Trim().ToLowerInvariant();
        q = q.Where(r => r.OrderCode.ToLowerInvariant().Contains(s)
                      || r.ProductName.ToLowerInvariant().Contains(s)
                      || r.StoreName.ToLowerInvariant().Contains(s)
                      || r.Sku.ToLowerInvariant().Contains(s));
    }
    int total = q.Count();
    int totalPages = Math.Max(1, (int)Math.Ceiling(total / (double)pageSize));
    if (page > totalPages) page = totalPages;
    var items = q.Skip((page - 1) * pageSize).Take(pageSize).ToList();
    return Results.Ok(new { total, page, pageSize, totalPages, records = items });
});

// ---------------- BƯỚC 2: LÀM SẠCH DỮ LIỆU ----------------
app.MapPost("/api/clean/run", (CleaningConfig? config) =>
{
    if (dataHub.RawData.Count == 0)
        return Results.BadRequest(new { message = "Chưa kết nối dữ liệu ở Bước 1. Vui lòng tải tệp dữ liệu." });

    var result = dataHub.RunDataCleaning(config ?? new CleaningConfig());
    return Results.Ok(result);
});

app.MapPut("/api/clean/raw-record/{index:int}", (int index, GenericDataRecord updated) =>
{
    var (ok, error) = dataHub.UpdateRawRecordValidated(index, updated);
    if (!ok)
        return Results.BadRequest(new { message = error });

    var result = dataHub.RunDataCleaning(dataHub.LastCleaningConfig, preserveVerification: false);
    return Results.Ok(new
    {
        message = "Đã lưu thay đổi, ghi nhật ký chỉnh sửa và cập nhật lại toàn bộ chỉ số!",
        cleanResult = result
    });
});

app.MapPost("/api/clean/raw-record/{index:int}/undo", (int index) =>
{
    var (ok, error) = dataHub.UndoRawRecord(index);
    if (!ok)
        return Results.BadRequest(new { message = error });

    var result = dataHub.RunDataCleaning(dataHub.LastCleaningConfig, preserveVerification: false);
    return Results.Ok(new
    {
        message = "Đã khôi phục giá trị trước đó cho bản ghi.",
        cleanResult = result
    });
});

app.MapGet("/api/clean/cleaned-records", (int page = 1, int pageSize = 50, string? search = null, string? sortBy = null, string? sortDir = "desc") =>
{
    page = Math.Max(1, page);
    pageSize = Math.Clamp(pageSize, 10, 200);
    IEnumerable<CleanedRecord> q = dataHub.CleanedData;
    if (!string.IsNullOrWhiteSpace(search))
    {
        var s = search.Trim().ToLowerInvariant();
        q = q.Where(r => r.OrderCode.ToLowerInvariant().Contains(s)
                      || r.ProductName.ToLowerInvariant().Contains(s)
                      || r.StoreName.ToLowerInvariant().Contains(s)
                      || r.Channel.ToLowerInvariant().Contains(s));
    }
    q = (sortBy ?? "").ToLowerInvariant() switch
    {
        "netrevenue" => sortDir == "asc" ? q.OrderBy(r => r.NetRevenue) : q.OrderByDescending(r => r.NetRevenue),
        "grossprofit" => sortDir == "asc" ? q.OrderBy(r => r.GrossProfit) : q.OrderByDescending(r => r.GrossProfit),
        "orderdate" => sortDir == "asc" ? q.OrderBy(r => r.OrderDate) : q.OrderByDescending(r => r.OrderDate),
        _ => q
    };
    int total = q.Count();
    int totalPages = Math.Max(1, (int)Math.Ceiling(total / (double)pageSize));
    if (page > totalPages) page = totalPages;
    var items = q.Skip((page - 1) * pageSize).Take(pageSize).ToList();
    return Results.Ok(new { total, page, pageSize, totalPages, records = items });
});

app.MapGet("/api/clean/exceptions", (int page = 1, int pageSize = 50, string? type = "ALL", string? search = null) =>
{
    page = Math.Max(1, page);
    pageSize = Math.Clamp(pageSize, 10, 200);
    IEnumerable<ExceptionRecord> q = dataHub.Exceptions;
    if (!string.IsNullOrWhiteSpace(type) && !type.Equals("ALL", StringComparison.OrdinalIgnoreCase))
    {
        q = q.Where(e => e.ExceptionType.Equals(type, StringComparison.OrdinalIgnoreCase));
    }
    if (!string.IsNullOrWhiteSpace(search))
    {
        var s = search.Trim().ToLowerInvariant();
        q = q.Where(e => e.Record.OrderCode.ToLowerInvariant().Contains(s)
                      || e.Record.StoreName.ToLowerInvariant().Contains(s)
                      || e.Record.ProductName.ToLowerInvariant().Contains(s)
                      || e.Reason.ToLowerInvariant().Contains(s));
    }
    int total = q.Count();
    int totalPages = Math.Max(1, (int)Math.Ceiling(total / (double)pageSize));
    if (page > totalPages) page = totalPages;
    var items = q.Skip((page - 1) * pageSize).Take(pageSize).ToList();
    var editedSet = dataHub.OriginalSnapshots.Keys.ToHashSet();
    var enriched = items.Select(e => new
    {
        e.RawIndex,
        e.ExceptionType,
        e.Reason,
        e.ActionTaken,
        e.Record,
        canUndo = editedSet.Contains(e.RawIndex)
    });
    return Results.Ok(new
    {
        total,
        page,
        pageSize,
        totalPages,
        records = enriched,
        countsByType = new
        {
            ALL = dataHub.Exceptions.Count,
            DUPLICATE = dataHub.Exceptions.Count(e => e.ExceptionType == "DUPLICATE"),
            CANCELLED = dataHub.Exceptions.Count(e => e.ExceptionType == "CANCELLED"),
            MISSING_OR_INVALID = dataHub.Exceptions.Count(e => e.ExceptionType == "MISSING_OR_INVALID")
        }
    });
});

app.MapGet("/api/clean/edit-logs", (int page = 1, int pageSize = 50) =>
{
    page = Math.Max(1, page);
    pageSize = Math.Clamp(pageSize, 10, 200);
    int total = dataHub.EditLogs.Count;
    var items = dataHub.EditLogs.AsEnumerable().Reverse().Skip((page - 1) * pageSize).Take(pageSize).ToList();
    return Results.Ok(new { total, page, pageSize, logs = items });
});

app.MapDelete("/api/clean/raw-record/{index:int}", (int index) =>
{
    if (!dataHub.DeleteRawRecord(index))
        return Results.NotFound(new { message = "Không tìm thấy bản ghi cần xóa." });

    var result = dataHub.RunDataCleaning(dataHub.LastCleaningConfig);
    return Results.Ok(new
    {
        message = "Đã loại bỏ bản ghi và cập nhật dữ liệu sạch.",
        cleanResult = result
    });
});

app.MapPost("/api/clean/verify", (VerifyPayload payload) =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu đã làm sạch để xác nhận." });

    dataHub.IsHumanVerified = payload.Approved;
    dataHub.VerificationNotes = payload.Notes ?? (payload.Approved ? "Chuyên viên đã kiểm tra và phê duyệt số liệu sạch." : "Dữ liệu bị từ chối.");

    return Results.Ok(new
    {
        approved = dataHub.IsHumanVerified,
        notes = dataHub.VerificationNotes,
        message = dataHub.IsHumanVerified
            ? "Phê duyệt dữ liệu sạch thành công! Đã mở khóa Bước 3: Phân tích dữ liệu."
            : "Đã từ chối dữ liệu. Vui lòng rà soát hoặc chỉnh sửa bản ghi ngoại lệ."
    });
});

// ---------------- BƯỚC 3: PHÂN TÍCH DỮ LIỆU ----------------
app.MapGet("/api/analyze/dashboard", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Vui lòng chạy Bước 2: Làm sạch dữ liệu trước khi phân tích." });

    var report = dataHub.PerformDeepAnalysis();
    return Results.Ok(report);
});

app.MapPost("/api/analyze/evaluate", (VerifyPayload payload) =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có kết quả phân tích để đánh giá." });

    dataHub.IsResultEvaluated = payload.Approved;
    return Results.Ok(new
    {
        evaluated = dataHub.IsResultEvaluated,
        message = dataHub.IsResultEvaluated
            ? "Đã đánh giá kết quả phân tích đạt chuẩn. Sẵn sàng xuất báo cáo."
            : "Đã đặt lại trạng thái chờ đánh giá."
    });
});

app.MapPost("/api/analyze/ask-copilot", (AskQuery query) =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để truy vấn." });

    var answer = dataHub.AnswerAnalystQuestion(query.Question ?? "");
    return Results.Ok(answer);
});

// ---------------- BƯỚC 4: XUẤT KẾT QUẢ ----------------
app.MapGet("/api/export/excel", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để tạo Excel BI Workbook." });

    var analysis = dataHub.PerformDeepAnalysis();
    var bytes = ExcelBiExporter.BuildWorkbook(dataHub, analysis);
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Ngoc_Linh_Data_Studio_BI_Workbook.xlsx");
});

app.MapGet("/api/export/word", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để tạo Báo cáo Word BI." });

    var analysis = dataHub.PerformDeepAnalysis();
    var bytes = WordBiExporter.BuildReport(dataHub, analysis);
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Ngoc_Linh_Data_Studio_BI_Report.docx");
});

app.MapGet("/api/export/html", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để tạo Interactive Dashboard HTML." });

    var analysis = dataHub.PerformDeepAnalysis();
    var html = HtmlDashboardExporter.BuildStandaloneHtml(dataHub, analysis);
    var bytes = Encoding.UTF8.GetBytes(html);
    return Results.File(bytes, "text/html; charset=utf-8", "Ngoc_Linh_Data_Studio_Dashboard.html");
});

app.MapGet("/api/export/html-preview", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để xem trước Dashboard HTML." });

    var analysis = dataHub.PerformDeepAnalysis();
    var html = HtmlDashboardExporter.BuildStandaloneHtml(dataHub, analysis);
    return Results.Ok(new { dataVersion = dataHub.DataVersion, html });
});

app.MapGet("/api/export/excel-preview", (string sheet = "CLEANED DATA", int page = 1, int pageSize = 50, string? search = null) =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để xem trước Excel." });

    page = Math.Max(1, page);
    pageSize = Math.Clamp(pageSize, 10, 100);
    var result = dataHub.GetExcelSheetPreview(sheet, page, pageSize, search);
    return Results.Ok(result);
});

app.MapGet("/api/export/preview", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu làm sạch để xem trước." });

    var analysis = dataHub.PerformDeepAnalysis();
    return Results.Ok(new
    {
        generatedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm"),
        processedAt = dataHub.ProcessedAt,
        dataVersion = dataHub.DataVersion,
        businessModel = dataHub.BusinessModel,
        editLogCount = dataHub.EditLogs.Count,
        sourceName = dataHub.SourceName,
        isVerified = dataHub.IsHumanVerified,
        isEvaluated = dataHub.IsResultEvaluated,
        excel = new
        {
            fileName = "Ngoc_Linh_Data_Studio_BI_Workbook.xlsx",
            sheetCount = 6,
            sheets = new[]
            {
                new { name = "RAW DATA", desc = $"Dữ liệu gốc ({dataHub.RawData.Count} bản ghi, 12 trường)" },
                new { name = "CLEANED DATA", desc = $"Excel Table dữ liệu sạch ({dataHub.CleanedData.Count} bản ghi, 18 trường)" },
                new { name = "DATA DICTIONARY", desc = "Từ điển 18 trường dữ liệu chuẩn BI" },
                new { name = "CLEANING LOG", desc = $"Nhật ký làm sạch ({dataHub.Exceptions.Count} ngoại lệ, {dataHub.CleaningLogs.Count} dòng log)" },
                new { name = "ANALYSIS", desc = "Bảng tổng hợp công thức động (Doanh thu, Lợi nhuận, Chi nhánh, Sản phẩm, Thời gian, Kênh)" },
                new { name = "DASHBOARD", desc = "5 KPI Cards, 4 Biểu đồ Excel thực tế & Insight tổng quan" }
            },
            cleanRecords = dataHub.CleanedData.Count,
            rawRecords = dataHub.RawData.Count
        },
        word = new
        {
            fileName = "Ngoc_Linh_Data_Studio_BI_Report.docx",
            sectionCount = 11,
            sections = new[]
            {
                "Phần 1. Trang bìa — Ngọc Linh Data Studio",
                "Phần 2. Tóm tắt điều hành (Executive Summary)",
                "Phần 3. Giới thiệu & Mục tiêu phân tích",
                "Phần 4. Tổng quan & Đánh giá chất lượng dữ liệu",
                "Phần 5. Phương pháp & Hệ thống chỉ số KPI",
                "Phần 6. Kết quả phân tích kinh doanh (Kèm bảng & 3 biểu đồ)",
                "Phần 7. Business Insights (Hiện tượng → Bằng chứng → Phân tích → Tác động → Vấn đề)",
                "Phần 8. Đề xuất giải pháp & Khuyến nghị (3 mức ưu tiên)",
                "Phần 9. Kế hoạch triển khai (Ngắn / Trung / Dài hạn)",
                "Phần 10. Kết luận hỗ trợ ra quyết định",
                "Phần 11. Phụ lục & Từ điển dữ liệu"
            }
        },
        analysis
    });
});

app.MapGet("/api/export/csv", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest("Chưa có dữ liệu làm sạch để xuất.");

    var csv = dataHub.ExportToCsv();
    var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(csv)).ToArray();
    return Results.File(bytes, "text/csv", $"NgocLinhDataStudio_CleanedData_{DateTime.Now:yyyyMMdd_HHmm}.csv");
});

app.MapGet("/api/export/cleaned-records", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu làm sạch." });

    return Results.Ok(dataHub.CleanedData);
});

app.MapGet("/api/export/summary-report", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu làm sạch để xuất báo cáo." });

    var report = dataHub.PerformDeepAnalysis();
    return Results.Ok(new
    {
        generatedDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"),
        analystAgency = "Ngọc Linh Data Studio - BI Data Cleansing & Analytics Workbench",
        sourceName = dataHub.SourceName,
        executiveSummary = report.ExecutiveSummary,
        financialMetrics = report.OverallMetrics,
        pricingAnalysis = report.PricingAnalysis,
        storePerformance = report.StorePerformance,
        productPerformance = report.ProductPerformance,
        channelBreakdown = report.ChannelBreakdown,
        timeSeries = report.TimeSeries,
        structuredRecommendations = report.StructuredRecommendations,
        strategicRecommendations = report.StrategicRecommendations
    });
});

app.Run();

// ================= DATA ANALYST HUB & PROCESSING ENGINE =================

public class BusinessModelPayload
{
    public string? BusinessModel { get; set; }
}

public class UploadBatchRequest
{
    public string? UploadId { get; set; }
    public string? FileName { get; set; }
    public string? ReplaceFileId { get; set; }
    public long FileSizeBytes { get; set; }
    public string? BusinessModel { get; set; }
    public List<string>? Headers { get; set; }
    public List<GenericDataRecord>? Records { get; set; }
    public int BatchIndex { get; set; }
    public int TotalBatches { get; set; }
}

public class ProcessingProgress
{
    public string Stage { get; set; } = "Sẵn sàng";
    public int Percent { get; set; }
    public int ProcessedRows { get; set; }
    public int TotalRows { get; set; }
    public long ElapsedMs { get; set; }
    public bool IsRunning { get; set; }
}

public class EditLogEntry
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];
    public int RawIndex { get; set; }
    public string OrderCode { get; set; } = "";
    public string Field { get; set; } = "";
    public string OldValue { get; set; } = "";
    public string NewValue { get; set; } = "";
    public string Action { get; set; } = "INLINE_EDIT";
    public string EditedAt { get; set; } = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
}

public class UploadParsedRequest
{
    public string? FileName { get; set; }
    public string? ReplaceFileId { get; set; }
    public List<string>? Headers { get; set; }
    public List<GenericDataRecord>? Records { get; set; }
}

public class CleaningConfig
{
    public bool RemoveDuplicates { get; set; } = true;
    public bool ExcludeCancelled { get; set; } = true;
    public bool DeductReturns { get; set; } = true;
    public bool ImputeMissingCogs { get; set; } = true;
    public bool NormalizeFormat { get; set; } = true;
}

public class VerifyPayload
{
    public bool Approved { get; set; } = true;
    public string? Notes { get; set; }
}

public class AskQuery
{
    public string? Question { get; set; }
}

public class GenericDataRecord
{
    public string OrderCode { get; set; } = "";
    public DateTime OrderDate { get; set; }
    public string CustomerName { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string StoreName { get; set; } = "";
    public string Sku { get; set; } = "";
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; } = 1;
    public decimal UnitPrice { get; set; }
    public decimal DiscountAmount { get; set; }
    public decimal UnitCost { get; set; }
    public string Status { get; set; } = "COMPLETED";
    public string Channel { get; set; } = "Store";
}

public class CleanedRecord : GenericDataRecord
{
    public decimal GrossRevenue { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal TotalCost { get; set; }
    public decimal GrossProfit { get; set; }
    public decimal MarginPercent { get; set; }
    public string CleanRemark { get; set; } = "Hợp lệ";
}

public class ExceptionRecord
{
    public int RawIndex { get; set; }
    public string ExceptionType { get; set; } = ""; // DUPLICATE, CANCELLED, RETURNED, MISSING_DATA, INVALID_FORMAT
    public string Reason { get; set; } = "";
    public string ActionTaken { get; set; } = "";
    public GenericDataRecord Record { get; set; } = new();
}

public class UploadedFileInfo
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];
    public string FileName { get; set; } = "";
    public int RecordCount { get; set; }
    public int ColumnCount { get; set; }
    public List<string> DetectedColumns { get; set; } = new();
    public List<string> MissingColumns { get; set; } = new();
    public string Status { get; set; } = "Hợp lệ";
    public long FileSizeBytes { get; set; }
    public DateTime UploadedAt { get; set; } = DateTime.Now;
    public List<GenericDataRecord> Records { get; set; } = new();

    public object ToSummary() => new
    {
        id = Id,
        fileName = FileName,
        recordCount = RecordCount,
        columnCount = ColumnCount,
        fileSizeBytes = FileSizeBytes,
        detectedColumns = DetectedColumns,
        missingColumns = MissingColumns,
        status = Status,
        uploadedAt = UploadedAt.ToString("dd/MM/yyyy HH:mm")
    };
}

public class DataAnalystHub
{
    public List<UploadedFileInfo> UploadedFiles { get; private set; } = new();
    public List<GenericDataRecord> RawData { get; private set; } = new();
    public List<CleanedRecord> CleanedData { get; private set; } = new();
    public List<ExceptionRecord> Exceptions { get; private set; } = new();
    public List<string> CleaningLogs { get; private set; } = new();
    public CleaningConfig LastCleaningConfig { get; private set; } = new();
    public CleaningStats LastStats { get; private set; } = new();
    public bool IsHumanVerified { get; set; } = false;
    public bool IsResultEvaluated { get; set; } = false;
    public string VerificationNotes { get; set; } = "";
    public string SourceName { get; set; } = "";
    public string BusinessModel { get; set; } = "Bán lẻ đa kênh";
    public int DataVersion { get; set; } = 1;
    public string ProcessedAt { get; set; } = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
    public ProcessingProgress Progress { get; private set; } = new();
    public List<EditLogEntry> EditLogs { get; private set; } = new();
    public Dictionary<int, GenericDataRecord> OriginalSnapshots { get; private set; } = new();
    private readonly Dictionary<string, (string FileName, List<string> Headers, List<GenericDataRecord> Records, long FileSizeBytes, string? ReplaceId)> _pendingBatches = new();

    private Dictionary<string, decimal> StandardCogs = new(StringComparer.OrdinalIgnoreCase);

    public void Reset()
    {
        UploadedFiles.Clear();
        RawData.Clear();
        CleanedData.Clear();
        Exceptions.Clear();
        CleaningLogs.Clear();
        EditLogs.Clear();
        OriginalSnapshots.Clear();
        _pendingBatches.Clear();
        LastStats = new();
        IsHumanVerified = false;
        IsResultEvaluated = false;
        VerificationNotes = "";
        SourceName = "";
        DataVersion++;
        ProcessedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        Progress = new();
    }

    public object GetPipelineStatus()
    {
        return new
        {
            step1_connected = RawData.Count > 0,
            step1_count = RawData.Count,
            step1_fileCount = UploadedFiles.Count,
            step1_source = SourceName,
            businessModel = BusinessModel,
            dataVersion = DataVersion,
            processedAt = ProcessedAt,
            step2_cleaned = CleanedData.Count > 0,
            step2_count = CleanedData.Count,
            step2_exceptions = Exceptions.Count,
            step2_verified = IsHumanVerified,
            step3_analyzed = CleanedData.Count > 0 && IsHumanVerified,
            step3_evaluated = IsResultEvaluated,
            step4_ready = CleanedData.Count > 0 && IsHumanVerified
        };
    }

    public List<object> GetFileSummaries() => UploadedFiles.Select(f => f.ToSummary()).ToList();

    private void RebuildRawFromFiles()
    {
        RawData = UploadedFiles.SelectMany(f => f.Records).ToList();
        SourceName = UploadedFiles.Count switch
        {
            0 => "",
            1 => UploadedFiles[0].FileName,
            _ => $"{UploadedFiles.Count} tệp ({string.Join(", ", UploadedFiles.Select(f => f.FileName))})"
        };
        CleanedData.Clear();
        Exceptions.Clear();
        CleaningLogs.Clear();
        LastStats = new();
        IsHumanVerified = false;
        IsResultEvaluated = false;
        EditLogs.Clear();
        OriginalSnapshots.Clear();
        DataVersion++;
        ProcessedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
    }

    public object AppendBatch(UploadBatchRequest req)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var uid = string.IsNullOrWhiteSpace(req.UploadId) ? Guid.NewGuid().ToString("N")[..8] : req.UploadId!;
        if (!string.IsNullOrWhiteSpace(req.BusinessModel)) BusinessModel = req.BusinessModel!;

        if (!_pendingBatches.TryGetValue(uid, out var state))
        {
            state = (req.FileName ?? "Uploaded_Data.xlsx", req.Headers ?? new List<string>(), new List<GenericDataRecord>(), req.FileSizeBytes, req.ReplaceFileId);
            _pendingBatches[uid] = state;
        }
        if (req.Records != null && req.Records.Count > 0)
            state.Records.AddRange(req.Records);

        int pct = req.TotalBatches > 0 ? Math.Min(95, (int)Math.Round(((req.BatchIndex + 1) / (double)req.TotalBatches) * 95)) : 95;
        Progress = new ProcessingProgress
        {
            Stage = req.BatchIndex + 1 >= req.TotalBatches ? "Đang kiểm tra cấu trúc" : "Đang đọc dữ liệu",
            Percent = pct,
            ProcessedRows = state.Records.Count,
            TotalRows = state.Records.Count,
            ElapsedMs = sw.ElapsedMilliseconds,
            IsRunning = req.BatchIndex + 1 < req.TotalBatches
        };

        if (req.BatchIndex + 1 < req.TotalBatches)
        {
            return new { completed = false, uploadId = uid, received = state.Records.Count, progress = Progress };
        }

        _pendingBatches.Remove(uid);
        var info = AddOrReplaceFile(state.FileName, state.Records, state.Headers, state.ReplaceId);
        info.FileSizeBytes = state.FileSizeBytes;
        Progress = new ProcessingProgress
        {
            Stage = "Hoàn tất đọc & kiểm tra cấu trúc",
            Percent = 100,
            ProcessedRows = info.RecordCount,
            TotalRows = info.RecordCount,
            ElapsedMs = sw.ElapsedMilliseconds,
            IsRunning = false
        };
        return new
        {
            completed = true,
            message = $"Đã kết nối thành công {info.RecordCount:N0} bản ghi từ '{info.FileName}'!",
            file = info.ToSummary(),
            files = GetFileSummaries(),
            summary = GetRawSummary(),
            preview = RawData.Take(50).ToList(),
            progress = Progress
        };
    }

    public (bool Success, UploadedFileInfo? Info, string Error) ParseAndAddFileStream(string fileName, Stream stream, long fileSizeBytes, string? replaceFileId = null)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            Progress = new ProcessingProgress { Stage = "Đang đọc dữ liệu", Percent = 5, IsRunning = true };
            var records = new List<GenericDataRecord>();
            var detectedHeaders = new List<string>();

            if (fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                using var sr = new StreamReader(stream, Encoding.UTF8);
                var content = sr.ReadToEnd();
                return ParseAndAddFile(fileName, content, replaceFileId);
            }

            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 65536);
            string? headerLine = reader.ReadLine();
            if (headerLine == null) return (false, null, "Tệp trống.");

            Progress.Stage = "Đang kiểm tra cấu trúc";
            Progress.Percent = 12;

            var rawHeaders = SplitCsvLine(headerLine).Select(h => h.Trim().Trim('\uFEFF').Trim('"')).ToList();
            detectedHeaders = rawHeaders;
            var headers = rawHeaders.Select(h => h.ToLowerInvariant()).ToList();

            int idxOrder = FindColIndex(headers, "ordercode", "order_code", "mã đơn", "mã đơn hàng", "mã gd", "ma gd", "mã giao dịch", "transaction", "order id", "orderid", "số hđ", "hóa đơn", "order", "code", "id");
            int idxDate = FindColIndex(headers, "orderdate", "order_date", "ngày gd", "ngày đơn", "ngày", "thời gian", "kỳ", "date", "time");
            int idxStore = FindColIndex(headers, "storename", "store", "chi nhánh", "đơn vị", "cửa hàng", "cơ sở", "branch", "shop", "khu vực");
            int idxSku = FindColIndex(headers, "sku", "mã sp", "mã sản phẩm", "productcode", "mã hàng");
            int idxName = FindColIndex(headers, "productname", "tên sản phẩm", "sản phẩm", "tên hàng", "mặt hàng", "dịch vụ", "product", "item");
            int idxQty = FindColIndex(headers, "quantity", "số lượng", "qty", "sl");
            int idxPrice = FindColIndex(headers, "unitprice", "đơn giá", "price", "giá bán", "doanh thu", "thành tiền", "tổng tiền", "giá trị", "revenue", "amount", "total", "sales");
            int idxDisc = FindColIndex(headers, "discountamount", "discount", "giảm giá", "chiết khấu");
            int idxCost = FindColIndex(headers, "unitcost", "giá vốn", "cogs", "cost");
            int idxStatus = FindColIndex(headers, "status", "trạng thái", "tình trạng");
            int idxCust = FindColIndex(headers, "customername", "khách hàng", "phụ trách", "customer");
            int idxChannel = FindColIndex(headers, "channel", "kênh", "kênh bán");

            string? line;
            int rowIdx = 1;
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                rowIdx++;
                var parts = SplitCsvLine(line);
                var storeName = GetVal(parts, idxStore, "Chưa phân loại");
                var qty = ParseInt(GetVal(parts, idxQty, "1"));
                var rawPrice = ParseDecimal(GetVal(parts, idxPrice, "0"));

                records.Add(new GenericDataRecord
                {
                    OrderCode = GetVal(parts, idxOrder, $"GD-{rowIdx - 1:D4}"),
                    StoreName = storeName,
                    StoreId = storeName,
                    Sku = GetVal(parts, idxSku, "N/A").ToUpperInvariant(),
                    ProductName = idxName >= 0 ? GetVal(parts, idxName, "Giao dịch chuẩn") : "Giao dịch chuẩn",
                    Quantity = qty,
                    UnitPrice = rawPrice,
                    DiscountAmount = ParseDecimal(GetVal(parts, idxDisc, "0")),
                    UnitCost = ParseDecimal(GetVal(parts, idxCost, "0")),
                    Status = NormalizeStatus(GetVal(parts, idxStatus, "COMPLETED")),
                    CustomerName = GetVal(parts, idxCust, ""),
                    Channel = GetVal(parts, idxChannel, "Khác"),
                    OrderDate = ParseDate(GetVal(parts, idxDate, ""))
                });

                if (rowIdx % 5000 == 0)
                {
                    Progress.Stage = "Đang đọc dữ liệu";
                    Progress.ProcessedRows = records.Count;
                    Progress.Percent = Math.Min(90, 15 + (records.Count / 1000));
                    Progress.ElapsedMs = sw.ElapsedMilliseconds;
                }
            }

            if (records.Count == 0) return (false, null, "Không nhận diện được dòng dữ liệu hợp lệ nào.");

            var info = AddOrReplaceFile(fileName, records, detectedHeaders, replaceFileId);
            info.FileSizeBytes = fileSizeBytes;
            Progress = new ProcessingProgress
            {
                Stage = "Hoàn tất đọc dữ liệu",
                Percent = 100,
                ProcessedRows = info.RecordCount,
                TotalRows = info.RecordCount,
                ElapsedMs = sw.ElapsedMilliseconds,
                IsRunning = false
            };
            return (true, info, "");
        }
        catch (Exception ex)
        {
            Progress = new ProcessingProgress { Stage = "Lỗi đọc dữ liệu", Percent = 0, IsRunning = false };
            return (false, null, ex.Message);
        }
    }

    public UploadedFileInfo AddOrReplaceFile(string fileName, List<GenericDataRecord> records, List<string> headers, string? replaceFileId = null)
    {
        // Chuẩn hóa sơ bộ các bản ghi
        foreach (var r in records)
        {
            if (r.OrderDate == default) r.OrderDate = DateTime.Today;
            r.Status = NormalizeStatus(r.Status);
            if (string.IsNullOrWhiteSpace(r.StoreName)) r.StoreName = "Chưa phân loại";
            if (string.IsNullOrWhiteSpace(r.Channel)) r.Channel = "Khác";
            if (string.IsNullOrWhiteSpace(r.Sku)) r.Sku = "N/A";
        }

        var required = new[] { "OrderCode", "ProductName", "Quantity", "UnitPrice" };
        var detected = headers.Count > 0 ? headers : new List<string> { "OrderCode", "OrderDate", "StoreName", "Sku", "ProductName", "Quantity", "UnitPrice", "DiscountAmount", "UnitCost", "Status", "Channel" };
        var missing = new List<string>();
        if (records.All(r => string.IsNullOrWhiteSpace(r.OrderCode))) missing.Add("Mã đơn");
        if (records.All(r => string.IsNullOrWhiteSpace(r.ProductName) || r.ProductName == "Không rõ tên")) missing.Add("Sản phẩm");
        if (records.All(r => r.UnitPrice == 0)) missing.Add("Đơn giá");

        var info = new UploadedFileInfo
        {
            FileName = fileName,
            RecordCount = records.Count,
            ColumnCount = detected.Count,
            DetectedColumns = detected,
            MissingColumns = missing,
            Status = missing.Count > 0 ? $"Thiếu: {string.Join(", ", missing)}" : "Cấu trúc chuẩn",
            Records = records
        };

        if (!string.IsNullOrWhiteSpace(replaceFileId))
        {
            var idx = UploadedFiles.FindIndex(f => f.Id == replaceFileId);
            if (idx >= 0) UploadedFiles[idx] = info;
            else UploadedFiles.Add(info);
        }
        else
        {
            UploadedFiles.Add(info);
        }

        RebuildRawFromFiles();
        return info;
    }

    public bool RemoveFile(string id)
    {
        var removed = UploadedFiles.RemoveAll(f => f.Id == id) > 0;
        if (removed) RebuildRawFromFiles();
        return removed;
    }

    public object GetRawSummary()
    {
        var stores = RawData.Select(r => r.StoreName).Distinct().Count();
        var skus = RawData.Select(r => r.Sku).Distinct().Count();
        var duplicates = RawData.GroupBy(r => $"{r.OrderCode}_{r.Sku}_{r.ProductName}".ToUpperInvariant()).Count(g => g.Count() > 1);
        var cancelled = RawData.Count(r => r.Status.Equals("CANCELLED", StringComparison.OrdinalIgnoreCase) || r.Status.Equals("FAILED", StringComparison.OrdinalIgnoreCase));
        var returns = RawData.Count(r => r.Status.Equals("RETURNED", StringComparison.OrdinalIgnoreCase));
        var missingOrInvalid = RawData.Count(r => r.Quantity <= 0 || r.UnitPrice <= 0 || string.IsNullOrWhiteSpace(r.OrderCode) || r.ProductName == "Không rõ tên");
        var grossValue = RawData.Sum(r => (r.Quantity * r.UnitPrice) - r.DiscountAmount);

        return new
        {
            totalRecords = RawData.Count,
            fileCount = UploadedFiles.Count,
            uniqueStores = stores,
            uniqueSkus = skus,
            duplicateGroups = duplicates,
            cancelledOrders = cancelled,
            returnedOrders = returns,
            missingOrInvalid = missingOrInvalid,
            estimatedRawGrossValue = grossValue,
            sourceName = SourceName
        };
    }

    public (bool Success, UploadedFileInfo? Info, string Error) ParseAndAddFile(string fileName, string content, string? replaceFileId = null)
    {
        try
        {
            var records = new List<GenericDataRecord>();
            var detectedHeaders = new List<string>();

            if (fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var parsed = JsonSerializer.Deserialize<List<GenericDataRecord>>(content, options);
                if (parsed != null && parsed.Count > 0)
                {
                    records = parsed;
                    detectedHeaders = new List<string> { "OrderCode", "OrderDate", "CustomerName", "StoreName", "Sku", "ProductName", "Quantity", "UnitPrice", "DiscountAmount", "UnitCost", "Status", "Channel" };
                }
            }
            else
            {
                using var reader = new StringReader(content);
                string? headerLine = reader.ReadLine();
                if (headerLine == null) return (false, null, "Tệp trống.");

                var rawHeaders = headerLine.Split(new[] { ',', ';', '\t' }).Select(h => h.Trim().Trim('\uFEFF').Trim('"')).ToList();
                detectedHeaders = rawHeaders;
                var headers = rawHeaders.Select(h => h.ToLower()).ToList();

                int idxOrder = FindColIndex(headers, "ordercode", "order_code", "mã đơn", "mã đơn hàng", "order id", "orderid", "order", "code");
                int idxDate = FindColIndex(headers, "orderdate", "order_date", "ngày", "ngày đơn", "date");
                int idxStore = FindColIndex(headers, "storename", "store", "chi nhánh", "cửa hàng");
                int idxSku = FindColIndex(headers, "sku", "mã sp", "mã sản phẩm", "productcode");
                int idxName = FindColIndex(headers, "productname", "tên sản phẩm", "sản phẩm", "product", "item");
                int idxQty = FindColIndex(headers, "quantity", "số lượng", "qty", "sl");
                int idxPrice = FindColIndex(headers, "unitprice", "đơn giá", "price", "giá bán");
                int idxDisc = FindColIndex(headers, "discountamount", "discount", "giảm giá", "chiết khấu");
                int idxCost = FindColIndex(headers, "unitcost", "giá vốn", "cogs", "cost");
                int idxStatus = FindColIndex(headers, "status", "trạng thái", "tình trạng");
                int idxCust = FindColIndex(headers, "customername", "khách hàng", "customer");
                int idxChannel = FindColIndex(headers, "channel", "kênh", "kênh bán");

                string? line;
                int rowIdx = 1;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    rowIdx++;
                    var parts = SplitCsvLine(line);
                    var storeName = GetVal(parts, idxStore, "Chưa phân loại");

                    records.Add(new GenericDataRecord
                    {
                        OrderCode = GetVal(parts, idxOrder, $"ROW-{rowIdx}"),
                        StoreName = storeName,
                        StoreId = storeName,
                        Sku = GetVal(parts, idxSku, "N/A").ToUpper(),
                        ProductName = GetVal(parts, idxName, "Không rõ tên"),
                        Quantity = ParseInt(GetVal(parts, idxQty, "1")),
                        UnitPrice = ParseDecimal(GetVal(parts, idxPrice, "0")),
                        DiscountAmount = ParseDecimal(GetVal(parts, idxDisc, "0")),
                        UnitCost = ParseDecimal(GetVal(parts, idxCost, "0")),
                        Status = NormalizeStatus(GetVal(parts, idxStatus, "COMPLETED")),
                        CustomerName = GetVal(parts, idxCust, ""),
                        Channel = GetVal(parts, idxChannel, "Khác"),
                        OrderDate = ParseDate(GetVal(parts, idxDate, ""))
                    });
                }
            }

            if (records.Count == 0) return (false, null, "Không nhận diện được dòng dữ liệu hợp lệ nào.");

            var info = AddOrReplaceFile(fileName, records, detectedHeaders, replaceFileId);
            return (true, info, "");
        }
        catch (Exception ex)
        {
            return (false, null, ex.Message);
        }
    }

    public (bool Ok, string Error) UpdateRawRecordValidated(int index, GenericDataRecord updated)
    {
        if (index < 0 || index >= RawData.Count)
            return (false, "Không tìm thấy bản ghi cần chỉnh sửa.");

        if (string.IsNullOrWhiteSpace(updated.OrderCode))
            return (false, "Mã đơn không được để trống.");
        if (string.IsNullOrWhiteSpace(updated.StoreName))
            return (false, "Đơn vị / Chi nhánh không được để trống.");
        if (updated.OrderDate == default)
            return (false, "Ngày giao dịch không hợp lệ.");

        var normStatus = NormalizeStatus(updated.Status);
        var cur = RawData[index];

        // Lưu bản gốc lần đầu để hỗ trợ Khôi phục (Undo) và không mất dữ liệu gốc
        if (!OriginalSnapshots.ContainsKey(index))
        {
            OriginalSnapshots[index] = new GenericDataRecord
            {
                OrderCode = cur.OrderCode,
                OrderDate = cur.OrderDate,
                CustomerName = cur.CustomerName,
                StoreId = cur.StoreId,
                StoreName = cur.StoreName,
                Sku = cur.Sku,
                ProductName = cur.ProductName,
                Quantity = cur.Quantity,
                UnitPrice = cur.UnitPrice,
                DiscountAmount = cur.DiscountAmount,
                UnitCost = cur.UnitCost,
                Status = cur.Status,
                Channel = cur.Channel
            };
        }

        void LogIfChanged(string field, string oldVal, string newVal)
        {
            if (oldVal != newVal)
            {
                EditLogs.Add(new EditLogEntry
                {
                    RawIndex = index,
                    OrderCode = updated.OrderCode.Trim(),
                    Field = field,
                    OldValue = oldVal,
                    NewValue = newVal,
                    Action = "INLINE_EDIT",
                    EditedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
                });
            }
        }

        LogIfChanged("Mã đơn", cur.OrderCode, updated.OrderCode.Trim());
        LogIfChanged("Chi nhánh", cur.StoreName, updated.StoreName.Trim());
        LogIfChanged("Ngày đơn", cur.OrderDate.ToString("yyyy-MM-dd"), updated.OrderDate.ToString("yyyy-MM-dd"));
        LogIfChanged("Trạng thái", cur.Status, normStatus);
        LogIfChanged("Tên sản phẩm", cur.ProductName, (updated.ProductName ?? "").Trim());
        LogIfChanged("Số lượng", cur.Quantity.ToString(), updated.Quantity.ToString());
        LogIfChanged("Đơn giá", cur.UnitPrice.ToString("N0"), updated.UnitPrice.ToString("N0"));

        updated.OrderCode = updated.OrderCode.Trim();
        updated.StoreName = updated.StoreName.Trim();
        updated.StoreId = updated.StoreName;
        updated.Status = normStatus;
        RawData[index] = updated;
        DataVersion++;
        ProcessedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        return (true, "");
    }

    public (bool Ok, string Error) UndoRawRecord(int index)
    {
        if (index < 0 || index >= RawData.Count)
            return (false, "Không tìm thấy bản ghi.");
        if (!OriginalSnapshots.TryGetValue(index, out var orig))
            return (false, "Bản ghi này chưa có lịch sử chỉnh sửa để khôi phục.");

        var cur = RawData[index];
        EditLogs.Add(new EditLogEntry
        {
            RawIndex = index,
            OrderCode = orig.OrderCode,
            Field = "Toàn bộ dòng",
            OldValue = $"{cur.OrderCode} | {cur.StoreName} | {cur.OrderDate:yyyy-MM-dd} | {cur.Status}",
            NewValue = $"{orig.OrderCode} | {orig.StoreName} | {orig.OrderDate:yyyy-MM-dd} | {orig.Status}",
            Action = "UNDO",
            EditedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
        });

        RawData[index] = orig;
        OriginalSnapshots.Remove(index);
        DataVersion++;
        ProcessedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        return (true, "");
    }

    public bool UpdateRawRecord(int index, GenericDataRecord updated)
    {
        return UpdateRawRecordValidated(index, updated).Ok;
    }

    public bool DeleteRawRecord(int index)
    {
        if (index < 0 || index >= RawData.Count) return false;
        RawData.RemoveAt(index);
        return true;
    }

    private int FindColIndex(List<string> headers, params string[] keywords)
    {
        foreach (var kw in keywords)
        {
            int exact = headers.FindIndex(h => h == kw);
            if (exact >= 0) return exact;
        }
        foreach (var kw in keywords)
        {
            int partial = headers.FindIndex(h => h.Contains(kw));
            if (partial >= 0) return partial;
        }
        return -1;
    }

    private static string NormalizeStatus(string? s)
    {
        var v = (s ?? "").Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(v)) return "COMPLETED";
        if (v.Contains("hoàn tất") || v.Contains("hoàn thành") || v.Contains("thành công") || v.Contains("complete") || v.Contains("success") || v.Contains("done")) return "COMPLETED";
        if (v.Contains("cancel") || v.Contains("hủy") || v.Contains("huỷ")) return "CANCELLED";
        if (v.Contains("fail") || v.Contains("lỗi") || v.Contains("thất bại")) return "FAILED";
        if (v.Contains("return") || v.Contains("hoàn") || v.Contains("trả")) return "RETURNED";
        return "COMPLETED";
    }

    private static DateTime ParseDate(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return DateTime.Today;
        string[] formats = { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM/dd/yyyy", "dd/MM/yyyy HH:mm", "yyyy-MM-dd HH:mm:ss" };
        if (DateTime.TryParseExact(s.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)) return d;
        return DateTime.TryParse(s, out d) ? d : DateTime.Today;
    }

    private string GetVal(List<string> parts, int index, string fallback)
    {
        if (index >= 0 && index < parts.Count && !string.IsNullOrWhiteSpace(parts[index]))
            return parts[index].Trim().Trim('"');
        return fallback;
    }

    private List<string> SplitCsvLine(string line)
    {
        var result = new List<string>();
        bool inQuotes = false;
        var sb = new StringBuilder();
        foreach (char c in line)
        {
            if (c == '\"') inQuotes = !inQuotes;
            else if ((c == ',' || c == ';' || c == '\t') && !inQuotes)
            {
                result.Add(sb.ToString());
                sb.Clear();
            }
            else sb.Append(c);
        }
        result.Add(sb.ToString());
        return result;
    }

    private int ParseInt(string s) => int.TryParse(s.Replace(".", "").Replace(",", ""), out int v) ? v : 1;
    private decimal ParseDecimal(string s) => decimal.TryParse(s.Replace(".", "").Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal v) ? v : 0;

    public object RunDataCleaning(CleaningConfig config, bool preserveVerification = false)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        LastCleaningConfig = config;
        CleaningLogs.Clear();
        CleanedData.Clear();
        Exceptions.Clear();
        if (!preserveVerification)
        {
            IsHumanVerified = false;
            IsResultEvaluated = false;
        }
        DataVersion++;
        ProcessedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        Progress = new ProcessingProgress { Stage = "Đang phát hiện dữ liệu trùng lặp", Percent = 20, TotalRows = RawData.Count, IsRunning = true };

        var seenKeys = new HashSet<string>();
        int duplicates = 0;
        int cancelled = 0;
        int returns = 0;
        int missingOrInvalid = 0;
        int imputed = 0;
        decimal preventedOvercount = 0;
        const int MaxDetailLogs = 300;

        CleaningLogs.Add($"[AGENT DATA CLEANER] Bắt đầu quét {RawData.Count:N0} bản ghi từ nguồn '{SourceName}'...");

        for (int idx = 0; idx < RawData.Count; idx++)
        {
            var r = RawData[idx];

            // Chuẩn hóa văn bản nếu bật
            if (config.NormalizeFormat)
            {
                r.OrderCode = (r.OrderCode ?? "").Trim().ToUpperInvariant();
                r.StoreName = (r.StoreName ?? "").Trim();
                r.ProductName = (r.ProductName ?? "").Trim();
                r.Sku = (r.Sku ?? "N/A").Trim().ToUpperInvariant();
                r.Channel = (r.Channel ?? "Khác").Trim();
                r.Status = NormalizeStatus(r.Status);
            }

            // Kiểm tra dữ liệu thiếu / sai định dạng
            if (r.Quantity <= 0 || r.UnitPrice <= 0 || string.IsNullOrWhiteSpace(r.OrderCode))
            {
                missingOrInvalid++;
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "MISSING_OR_INVALID",
                    Reason = r.Quantity <= 0 ? "Số lượng <= 0" : r.UnitPrice <= 0 ? "Thiếu đơn giá / doanh thu (<= 0)" : "Thiếu mã đơn / mã giao dịch",
                    ActionTaken = "Tạm giữ để kiểm tra",
                    Record = r
                });
                if (CleaningLogs.Count < MaxDetailLogs)
                    CleaningLogs.Add($"[NGOẠI LỆ] Dòng #{idx + 1} ({r.OrderCode}): thiếu dữ liệu hoặc sai định dạng (SL={r.Quantity}, Giá={r.UnitPrice:N0}).");
                continue;
            }

            var key = $"{r.OrderCode}_{r.Sku}_{r.ProductName}".ToUpperInvariant();

            // 1. Khử trùng
            if (config.RemoveDuplicates && seenKeys.Contains(key))
            {
                duplicates++;
                var dupVal = (r.Quantity * r.UnitPrice) - r.DiscountAmount;
                preventedOvercount += dupVal;
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "DUPLICATE",
                    Reason = $"Trùng mã đơn & sản phẩm ({r.OrderCode})",
                    ActionTaken = "Đã loại bỏ bản ghi trùng",
                    Record = r
                });
                if (CleaningLogs.Count < MaxDetailLogs)
                    CleaningLogs.Add($"[DEDUPLICATE] Loại bỏ bản ghi trùng lặp: Đơn {r.OrderCode} | {r.ProductName} (Ngăn chặn {dupVal:N0} đ).");
                continue;
            }
            seenKeys.Add(key);

            // 2. Loại đơn hủy / thất bại
            if (config.ExcludeCancelled && (r.Status.Equals("CANCELLED", StringComparison.OrdinalIgnoreCase) || r.Status.Equals("FAILED", StringComparison.OrdinalIgnoreCase)))
            {
                cancelled++;
                var cancelVal = (r.Quantity * r.UnitPrice) - r.DiscountAmount;
                preventedOvercount += cancelVal;
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "CANCELLED",
                    Reason = $"Trạng thái đơn: {r.Status}",
                    ActionTaken = "Loại khỏi doanh thu thuần",
                    Record = r
                });
                if (CleaningLogs.Count < MaxDetailLogs)
                    CleaningLogs.Add($"[EXCLUDE CANCELLED] Đơn {r.OrderCode} ({r.Status}) → loại khỏi doanh thu ({cancelVal:N0} đ).");
                continue;
            }

            // 3. Giá vốn thực tế (COGS) - KHÔNG tự bịa nếu dữ liệu nguồn không có
            decimal unitCost = r.UnitCost > 0 ? r.UnitCost : 0m;
            if (unitCost <= 0)
            {
                imputed++; // Đếm số dòng thiếu giá vốn thực tế
            }

            decimal gross = r.Quantity * r.UnitPrice;
            decimal net = gross - r.DiscountAmount;
            decimal totalCost = unitCost > 0 ? r.Quantity * unitCost : 0m;

            // 4. Xử lý đổi trả
            string remark = unitCost <= 0 ? "Thiếu giá vốn thực" : "Hợp lệ";
            if (r.Status.Equals("RETURNED", StringComparison.OrdinalIgnoreCase))
            {
                returns++;
                remark = "Hàng hoàn (Giảm trừ)";
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "RETURNED",
                    Reason = "Đơn hàng hoàn trả",
                    ActionTaken = config.DeductReturns ? "Đã giảm trừ DT & hoàn giá vốn" : "Giữ nguyên chờ đối soát",
                    Record = r
                });
                if (config.DeductReturns)
                {
                    gross = -gross;
                    net = -net;
                    totalCost = -totalCost;
                    if (CleaningLogs.Count < MaxDetailLogs)
                        CleaningLogs.Add($"[RETURN ADJUSTMENT] Đơn {r.OrderCode} là hàng trả lại → giảm trừ doanh thu & hoàn nhập giá vốn.");
                }
            }

            decimal grossProfit = unitCost > 0 ? (net - totalCost) : 0m;
            decimal marginPct = (unitCost > 0 && net > 0) ? Math.Round((grossProfit / net) * 100, 2) : 0;

            CleanedData.Add(new CleanedRecord
            {
                OrderCode = r.OrderCode,
                OrderDate = r.OrderDate,
                CustomerName = r.CustomerName,
                StoreId = r.StoreId,
                StoreName = r.StoreName,
                Sku = r.Sku,
                ProductName = r.ProductName,
                Quantity = r.Quantity,
                UnitPrice = r.UnitPrice,
                DiscountAmount = r.DiscountAmount,
                UnitCost = unitCost,
                GrossRevenue = gross,
                NetRevenue = net,
                TotalCost = totalCost,
                GrossProfit = grossProfit,
                MarginPercent = marginPct,
                Status = r.Status,
                Channel = r.Channel,
                CleanRemark = remark
            });
        }

        if (imputed > 0)
            CleaningLogs.Add($"[LƯU Ý GIÁ VỐN] {imputed} dòng chưa có dữ liệu giá vốn.");

        decimal gmvVal = CleanedData.Where(x => x.Status != "RETURNED").Sum(x => x.GrossRevenue);
        decimal actualRevVal = CleanedData.Sum(x => x.NetRevenue);

        CleaningLogs.Add($"[AGENT COMPLETE] Hoàn tất làm sạch! {RawData.Count} dòng thô -> {CleanedData.Count} dòng sạch. Doanh thu thực tế: {actualRevVal:N0} đ (GMV: {gmvVal:N0} đ).");
        CleaningLogs.Add($"[STATISTICS] Trùng: {duplicates} · Hủy/lỗi: {cancelled} · Hoàn trả: {returns} · Dòng trống/sai ĐD: {missingOrInvalid} · Sai lệch ngăn ngừa: {preventedOvercount:N0} đ.");

        LastStats = new CleaningStats
        {
            RawCount = RawData.Count,
            CleanedCount = CleanedData.Count,
            DuplicatesRemoved = duplicates,
            CancelledExcluded = cancelled,
            ReturnsAdjusted = returns,
            MissingOrInvalid = missingOrInvalid,
            ImputedCount = imputed,
            PreventedOvercountValue = preventedOvercount
        };

        Progress = new ProcessingProgress
        {
            Stage = "Hoàn tất làm sạch & chuẩn hóa dữ liệu",
            Percent = 100,
            ProcessedRows = RawData.Count,
            TotalRows = RawData.Count,
            ElapsedMs = sw.ElapsedMilliseconds,
            IsRunning = false
        };

        return new
        {
            success = true,
            dataVersion = DataVersion,
            processedAt = ProcessedAt,
            elapsedMs = sw.ElapsedMilliseconds,
            rawCount = RawData.Count,
            cleanedCount = CleanedData.Count,
            duplicatesRemoved = duplicates,
            cancelledExcluded = cancelled,
            returnsAdjusted = returns,
            missingOrInvalid = missingOrInvalid,
            missingCostCount = imputed,
            gmv = gmvVal,
            actualRevenue = actualRevVal,
            preventedOvercountValue = preventedOvercount,
            logs = CleaningLogs,
            exceptionCount = Exceptions.Count,
            separatedDuplicates = Exceptions.Where(e => e.ExceptionType == "DUPLICATE").Take(50).ToList(),
            exceptions = Exceptions.Take(50).ToList(),
            sampleCleaned = CleanedData.Take(50).ToList()
        };
    }

    public object GetExcelSheetPreview(string sheet, int page, int pageSize, string? search)
    {
        var s = (search ?? "").Trim().ToLowerInvariant();
        string norm = (sheet ?? "CLEANED DATA").Trim().ToUpperInvariant();

        if (norm == "RAW DATA")
        {
            IEnumerable<GenericDataRecord> q = RawData;
            if (!string.IsNullOrEmpty(s))
                q = q.Where(r => r.OrderCode.ToLowerInvariant().Contains(s) || r.StoreName.ToLowerInvariant().Contains(s) || r.ProductName.ToLowerInvariant().Contains(s));
            int total = q.Count();
            int totalPages = Math.Max(1, (int)Math.Ceiling(total / (double)pageSize));
            if (page > totalPages) page = totalPages;
            var rows = q.Skip((page - 1) * pageSize).Take(pageSize).Select(r => new string[]
            {
                r.OrderCode, r.OrderDate == default ? "" : r.OrderDate.ToString("yyyy-MM-dd"), r.CustomerName, r.StoreName,
                r.Sku, r.ProductName, r.Quantity.ToString("N0"), r.UnitPrice.ToString("N0"), r.DiscountAmount.ToString("N0"),
                r.UnitCost.ToString("N0"), r.Status, r.Channel
            }).ToList();
            return new
            {
                sheet = "RAW DATA",
                headers = new[] { "Mã đơn", "Ngày đơn", "Khách hàng", "Chi nhánh", "SKU", "Tên sản phẩm", "SL", "Đơn giá", "Chiết khấu", "Giá vốn ĐV", "Trạng thái", "Kênh" },
                total, page, pageSize, totalPages, rows
            };
        }

        if (norm == "DATA DICTIONARY")
        {
            var dict = new[]
            {
                new[] { "OrderCode", "Mã đơn", "Văn bản", "Khóa chính kiểm tra trùng lặp & đếm đơn" },
                new[] { "OrderDate", "Ngày đơn", "Ngày tháng (YYYY-MM-DD)", "Phân tích xu hướng tăng trưởng theo thời gian" },
                new[] { "CustomerName", "Khách hàng", "Văn bản", "Phân tích tập khách hàng (ẩn khỏi HTML công khai)" },
                new[] { "StoreName", "Chi nhánh", "Văn bản", "Phân tích hiệu quả theo đơn vị / chi nhánh" },
                new[] { "Sku", "Mã SKU", "Văn bản", "Kết hợp mã đơn để khử trùng & nhóm sản phẩm" },
                new[] { "ProductName", "Tên sản phẩm", "Văn bản", "Xếp hạng Top sản phẩm theo doanh thu & lợi nhuận" },
                new[] { "Quantity", "Số lượng", "Số nguyên", "Tính sản lượng & doanh thu gộp" },
                new[] { "UnitPrice", "Đơn giá", "VNĐ", "Phân tích chính sách giá bán" },
                new[] { "DiscountAmount", "Chiết khấu", "VNĐ", "Đánh giá tỷ lệ giảm giá bình quân" },
                new[] { "UnitCost", "Giá vốn ĐV", "VNĐ", "Tính giá vốn & lợi nhuận gộp" },
                new[] { "NetRevenue", "Doanh thu thuần", "Công thức BI", "Doanh thu thực tế sau chiết khấu & hoàn trả" },
                new[] { "CleanRemark", "Ghi chú làm sạch", "Nhãn trạng thái", "Truy vết minh bạch dữ liệu ngoại lệ" }
            };
            return new
            {
                sheet = "DATA DICTIONARY",
                headers = new[] { "Trường chuẩn", "Tên hiển thị", "Kiểu / Định dạng", "Vai trò phân tích BI" },
                total = dict.Length, page = 1, pageSize = dict.Length, totalPages = 1, rows = dict
            };
        }

        if (norm == "CLEANING LOG")
        {
            var logs = new List<string[]>
            {
                new[] { "Tổng bản ghi đầu vào", LastStats.RawCount.ToString("N0"), "Tiếp nhận nguyên trạng" },
                new[] { "Bản ghi trùng lặp đã tách", LastStats.DuplicatesRemoved.ToString("N0"), "Khử trùng theo Mã đơn + SKU + Sản phẩm" },
                new[] { "Đơn hủy / lỗi loại trừ", LastStats.CancelledExcluded.ToString("N0"), "Loại khỏi doanh thu thực tế" },
                new[] { "Đơn hoàn trả đối soát", LastStats.ReturnsAdjusted.ToString("N0"), "Giảm trừ doanh thu thực tế" },
                new[] { "Dòng thiếu / sai định dạng", LastStats.MissingOrInvalid.ToString("N0"), "Đưa vào bảng ngoại lệ" },
                new[] { "Dòng thiếu giá vốn thực", LastStats.ImputedCount.ToString("N0"), "Để trống (0 đ)" },
                new[] { "Số lần chỉnh sửa trực tiếp", EditLogs.Count.ToString("N0"), "Ghi nhận nhật ký có thể khôi phục" },
                new[] { "Bản ghi hợp lệ cuối cùng", LastStats.CleanedCount.ToString("N0"), IsHumanVerified ? "Đã kiểm duyệt" : "Chờ duyệt" }
            };
            return new
            {
                sheet = "CLEANING LOG",
                headers = new[] { "Chỉ tiêu kiểm soát", "Giá trị", "Quy tắc áp dụng" },
                total = logs.Count, page = 1, pageSize = logs.Count, totalPages = 1, rows = logs
            };
        }

        if (norm == "ANALYSIS" || norm == "DASHBOARD")
        {
            var a = PerformDeepAnalysis();
            var m = a.OverallMetrics;
            var p = a.PricingAnalysis;
            var rows = new List<string[]>
            {
                new[] { "Tổng doanh thu gộp (GMV)", $"{m.TotalGrossRevenue:N0} đ", "Tổng tiền hàng trước chiết khấu & hoàn" },
                new[] { "Tổng chiết khấu", $"{m.TotalDiscountAmount:N0} đ", $"Tỷ lệ CK bình quân: {p.AverageDiscountRate}%" },
                new[] { "Doanh thu thuần thực tế", $"{m.TotalNetRevenue:N0} đ", "Đã trừ đơn hủy, trùng lặp & hoàn trả" },
                new[] { "Tổng giá vốn (COGS)", m.HasCostData ? $"{m.TotalCost:N0} đ" : "—", m.HasCostData ? "Từ dữ liệu giá vốn" : "" },
                new[] { "Lợi nhuận gộp", m.HasCostData ? $"{m.TotalGrossProfit:N0} đ" : "Chưa có dữ liệu giá vốn", $"Biên LN: {(m.HasCostData ? m.GrossMarginPercent + "%" : "—")}" },
                new[] { "Giá trị đơn trung bình (AOV)", $"{p.AverageOrderValue:N0} đ", $"Trên {m.TotalOrders:N0} bản ghi hợp lệ" },
                new[] { "Tỷ lệ hoàn trả", $"{m.ReturnRatePercent}%", $"{m.ReturnOrders:N0} đơn hoàn trả" },
                new[] { "Tỷ lệ hủy đơn", $"{m.CancelRatePercent}%", $"{LastStats.CancelledExcluded:N0} đơn hủy/lỗi" },
                new[] { "Doanh thu ảo đã ngăn chặn", $"{m.PreventedOvercountValue:N0} đ", "Từ đơn trùng lặp và đơn hủy" }
            };
            return new
            {
                sheet = norm,
                headers = new[] { "Chỉ số BI", "Giá trị thực tế", "Diễn giải" },
                total = rows.Count, page = 1, pageSize = rows.Count, totalPages = 1, rows
            };
        }

        // Default: CLEANED DATA
        IEnumerable<CleanedRecord> cq = CleanedData;
        if (!string.IsNullOrEmpty(s))
            cq = cq.Where(r => r.OrderCode.ToLowerInvariant().Contains(s) || r.StoreName.ToLowerInvariant().Contains(s) || r.ProductName.ToLowerInvariant().Contains(s));
        int cTotal = cq.Count();
        int cTotalPages = Math.Max(1, (int)Math.Ceiling(cTotal / (double)pageSize));
        if (page > cTotalPages) page = cTotalPages;
        var cRows = cq.Skip((page - 1) * pageSize).Take(pageSize).Select(r => new string[]
        {
            r.OrderCode, r.OrderDate.ToString("yyyy-MM-dd"), r.StoreName, r.Sku, r.ProductName,
            r.Quantity.ToString("N0"), r.UnitPrice.ToString("N0"), r.DiscountAmount.ToString("N0"),
            r.NetRevenue.ToString("N0"), r.TotalCost.ToString("N0"), r.GrossProfit.ToString("N0"),
            r.Status, r.Channel, r.CleanRemark
        }).ToList();
        return new
        {
            sheet = "CLEANED DATA",
            headers = new[] { "Mã đơn", "Ngày đơn", "Chi nhánh", "SKU", "Tên sản phẩm", "SL", "Đơn giá", "Chiết khấu", "Doanh thu thuần", "Giá vốn", "Lợi nhuận gộp", "Trạng thái", "Kênh", "Ghi chú" },
            total = cTotal, page, pageSize, totalPages = cTotalPages, rows = cRows
        };
    }

    public DataSchemaProfile DetectSchemaProfile()
    {
        var cols = UploadedFiles.SelectMany(f => f.DetectedColumns).Distinct().ToList();
        var valid = CleanedData.Count > 0 ? CleanedData.Cast<GenericDataRecord>().ToList() : RawData;
        return new DataSchemaProfile
        {
            HasDate = valid.Any(x => x.OrderDate != default) && valid.Select(x => x.OrderDate.Date).Distinct().Count() >= 1,
            HasRevenue = valid.Any(x => x.UnitPrice > 0),
            HasCost = valid.Any(x => x.UnitCost > 0),
            HasProduct = valid.Any(x => !string.IsNullOrWhiteSpace(x.ProductName) && x.ProductName != "Giao dịch chuẩn" && x.ProductName != "Không rõ tên"),
            HasStore = valid.Any(x => !string.IsNullOrWhiteSpace(x.StoreName) && x.StoreName != "Chưa phân loại"),
            HasCustomer = valid.Any(x => !string.IsNullOrWhiteSpace(x.CustomerName)),
            HasChannel = valid.Any(x => !string.IsNullOrWhiteSpace(x.Channel) && x.Channel != "Khác"),
            HasDiscount = valid.Any(x => x.DiscountAmount > 0),
            HasStatus = valid.Any(x => !string.IsNullOrWhiteSpace(x.Status)),
            DetectedColumns = cols
        };
    }

    public DeepAnalysisResult PerformDeepAnalysis()
    {
        var valid = CleanedData;
        bool hasCostData = valid.Any(x => x.UnitCost > 0);
        decimal totalGross = valid.Sum(x => x.GrossRevenue > 0 ? x.GrossRevenue : 0);
        decimal totalDiscount = valid.Sum(x => x.Status == "RETURNED" ? 0 : x.DiscountAmount);
        decimal totalNet = valid.Sum(x => x.NetRevenue);
        decimal totalCost = hasCostData ? valid.Sum(x => x.TotalCost) : 0m;
        decimal totalProfit = hasCostData ? valid.Where(x => x.UnitCost > 0).Sum(x => x.GrossProfit) : 0m;
        decimal netWithCost = valid.Where(x => x.UnitCost > 0).Sum(x => x.NetRevenue);
        decimal grossMargin = (hasCostData && netWithCost > 0) ? Math.Round((totalProfit / netWithCost) * 100, 2) : 0m;

        int totalOrders = valid.Count;
        int returnOrders = valid.Count(x => x.Status == "RETURNED");
        decimal returnRate = totalOrders > 0 ? Math.Round(((decimal)returnOrders / totalOrders) * 100, 2) : 0;
        decimal cancelRate = LastStats.RawCount > 0 ? Math.Round(((decimal)LastStats.CancelledExcluded / LastStats.RawCount) * 100, 2) : 0;

        // Phân tích giá bán & chiết khấu
        var nonReturnOrders = valid.Where(x => x.Status != "RETURNED").ToList();
        decimal avgOrderValue = nonReturnOrders.Count > 0 ? Math.Round(totalNet / nonReturnOrders.Count, 0) : 0;
        decimal avgDiscountRate = totalGross > 0 ? Math.Round((totalDiscount / totalGross) * 100, 2) : 0;
        int discountedOrderCount = nonReturnOrders.Count(x => x.DiscountAmount > 0);
        decimal maxDiscountOrderPct = nonReturnOrders.Where(x => x.GrossRevenue > 0)
            .Select(x => Math.Round((x.DiscountAmount / x.GrossRevenue) * 100, 2))
            .DefaultIfEmpty(0).Max();

        // Phân tích theo Chi nhánh
        var storeResults = valid.GroupBy(x => x.StoreName).Select(g =>
        {
            decimal net = g.Sum(x => x.NetRevenue);
            decimal cost = g.Sum(x => x.TotalCost);
            decimal profit = net - cost;
            decimal margin = net > 0 ? Math.Round((profit / net) * 100, 2) : 0;
            int rets = g.Count(x => x.Status == "RETURNED");
            decimal retRate = g.Count() > 0 ? Math.Round(((decimal)rets / g.Count()) * 100, 2) : 0;

            string storeStatus = "Vận hành tốt";
            if (margin < 12) storeStatus = "Biên lợi nhuận mỏng";
            else if (retRate > 15) storeStatus = "Cảnh báo đổi trả cao";

            return new StoreAnalysisItem
            {
                StoreName = g.Key,
                OrderCount = g.Count(),
                NetRevenue = net,
                GrossProfit = profit,
                MarginPercent = margin,
                ReturnCount = rets,
                ReturnRatePercent = retRate,
                StatusAssessment = storeStatus
            };
        }).OrderByDescending(s => s.NetRevenue).ToList();

        // Phân tích theo Sản phẩm
        var productResults = valid
            .GroupBy(x => new { x.Sku, x.ProductName })
            .Select(g =>
            {
                decimal net = g.Sum(x => x.NetRevenue);
                decimal cost = g.Sum(x => x.TotalCost);
                decimal profit = net - cost;
                decimal disc = g.Sum(x => x.Status == "RETURNED" ? 0 : x.DiscountAmount);
                return new ProductAnalysisItem
                {
                    Sku = g.Key.Sku,
                    ProductName = g.Key.ProductName,
                    QuantitySold = g.Sum(x => x.Status == "RETURNED" ? -x.Quantity : x.Quantity),
                    NetRevenue = net,
                    GrossProfit = profit,
                    DiscountAmount = disc,
                    MarginPercent = net > 0 ? Math.Round((profit / net) * 100, 2) : 0
                };
            })
            .OrderByDescending(p => p.NetRevenue)
            .ToList();

        // Phân tích Kênh bán
        var channelResults = valid
            .GroupBy(x => x.Channel)
            .Select(g => new ChannelAnalysisItem
            {
                ChannelName = g.Key,
                OrderCount = g.Count(),
                NetRevenue = g.Sum(x => x.NetRevenue),
                RevenueSharePercent = totalNet > 0 ? Math.Round((g.Sum(x => x.NetRevenue) / totalNet) * 100, 1) : 0
            })
            .OrderByDescending(c => c.NetRevenue)
            .ToList();

        // Phân tích tăng trưởng theo thời gian
        var timeSeries = valid
            .GroupBy(x => x.OrderDate.Date)
            .OrderBy(g => g.Key)
            .Select(g => new TimeSeriesPoint
            {
                Date = g.Key.ToString("yyyy-MM-dd"),
                Label = g.Key.ToString("dd/MM"),
                NetRevenue = g.Sum(x => x.NetRevenue),
                GrossProfit = g.Sum(x => x.GrossProfit),
                Orders = g.Count()
            })
            .ToList();

        decimal growthRate = 0;
        if (timeSeries.Count >= 2 && timeSeries.First().NetRevenue > 0)
        {
            growthRate = Math.Round(((timeSeries.Last().NetRevenue - timeSeries.First().NetRevenue) / timeSeries.First().NetRevenue) * 100, 2);
        }

        var schema = DetectSchemaProfile();

        // 6 khối Insight quản trị cốt lõi (tự động căn chỉnh theo cấu trúc & số liệu từng file dữ liệu thô)
        string today = DateTime.Now.ToString("dd/MM/yyyy");
        var structuredRecs = new List<StructuredRecommendation>();
        int rawTotal = Math.Max(1, LastStats.RawCount > 0 ? LastStats.RawCount : totalOrders);
        int cleanPct = (int)Math.Round((totalOrders / (double)rawTotal) * 100);
        var bestStore = schema.HasStore ? storeResults.FirstOrDefault() : null;
        decimal bestStoreShare = (bestStore != null && totalNet > 0) ? Math.Round((bestStore.NetRevenue / totalNet) * 100, 1) : 100m;
        var worstRetStore = schema.HasStore ? (storeResults.OrderByDescending(s => s.ReturnRatePercent).FirstOrDefault() ?? bestStore) : null;
        var lowRevStore = (schema.HasStore && storeResults.Count > 1) ? storeResults.OrderBy(s => s.NetRevenue).FirstOrDefault() : null;
        var topProd = schema.HasProduct ? productResults.FirstOrDefault() : null;
        var lowProd = (schema.HasProduct && productResults.Count > 1) ? productResults.OrderBy(p => p.NetRevenue).FirstOrDefault() : null;
        decimal topProdShare = (topProd != null && totalNet > 0) ? Math.Round((topProd.NetRevenue / totalNet) * 100, 1) : 100m;
        var topChan = schema.HasChannel ? channelResults.FirstOrDefault() : null;
        decimal costSharePct = (hasCostData && totalNet > 0) ? Math.Round((totalCost / totalNet) * 100, 1) : 0m;

        // INSIGHT 01 · QUY MÔ GIÁ TRỊ
        string leadEntity = bestStore != null ? $"đơn vị {bestStore.StoreName} ({bestStoreShare}%)"
            : topProd != null ? $"sản phẩm {topProd.ProductName} ({topProdShare}%)"
            : topChan != null ? $"kênh {topChan.ChannelName} ({topChan.RevenueSharePercent}%)"
            : $"toàn bộ {totalOrders:N0} giao dịch";
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "INSIGHT 01 · QUY MÔ GIÁ TRỊ",
            Title = "Đóng Góp Doanh Thu & Trọng Tâm Vận Hành",
            Content = $"Từ nguồn '{SourceName}', tổng doanh thu thuần thực tế đạt {totalNet:N0} đ (GMV {totalGross:N0} đ). Trọng tâm đóng góp lớn nhất thuộc về {leadEntity}.",
            DetectedDate = today,
            Priority = "Trung bình"
        });

        // INSIGHT 02 · CHẤT LƯỢNG DỮ LIỆU
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "INSIGHT 02 · CHẤT LƯỢNG DỮ LIỆU",
            Title = "Độ Sạch Đầu Vào & Mức Độ Ổn Định",
            Content = $"Đã chuẩn hóa {totalOrders:N0} / {rawTotal:N0} bản ghi ({cleanPct}% dữ liệu sạch), loại bỏ {LastStats.DuplicatesRemoved + LastStats.CancelledExcluded:N0} dòng trùng/hủy (ngăn chặn {LastStats.PreventedOvercountValue:N0} đ). Tỷ lệ hoàn trả thực tế: {returnRate}%.",
            DetectedDate = today,
            Priority = LastStats.DuplicatesRemoved + LastStats.CancelledExcluded > 0 ? "Trung bình" : "Thấp"
        });

        // INSIGHT 03 · PHÂN TÍCH ĐỘ LỆCH
        string devDesc = lowRevStore != null && worstRetStore != null
            ? $"Đơn vị quy mô thấp nhất là {lowRevStore.StoreName} ({lowRevStore.NetRevenue:N0} đ); đơn vị có tỷ lệ hoàn trả cao nhất là {worstRetStore.StoreName} ({worstRetStore.ReturnRatePercent}%)."
            : lowProd != null
                ? $"Sản phẩm có doanh thu thấp nhất là {lowProd.ProductName} ({lowProd.NetRevenue:N0} đ); tỷ lệ hoàn toàn hệ thống ở mức {returnRate}%, hủy đơn {cancelRate}%."
                : $"Tỷ lệ hoàn trả toàn tệp ở mức {returnRate}% ({returnOrders} đơn), tỷ lệ hủy đơn {cancelRate}% ({LastStats.CancelledExcluded} đơn).";
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "INSIGHT 03 · PHÂN TÍCH ĐỘ LỆCH",
            Title = "Điểm Nghẽn Vận Hành & Tỷ Lệ Hoàn Trả",
            Content = devDesc,
            DetectedDate = today,
            Priority = returnRate > 5 || (worstRetStore?.ReturnRatePercent ?? 0) > 10 ? "Cao" : "Trung bình"
        });

        // INSIGHT 04 · BIÊN HIỆU QUẢ
        string effDesc = hasCostData
            ? $"Lợi nhuận gộp đạt {totalProfit:N0} đ (biên LN {grossMargin}%, tỷ trọng giá vốn {costSharePct}%). {(topProd != null ? $"Sản phẩm chủ lực {topProd.ProductName} chiếm {topProdShare}% cơ cấu." : $"Giá trị đơn TB (AOV) đạt {avgOrderValue:N0} đ.")}"
            : $"{(topProd != null ? $"Sản phẩm chủ lực {topProd.ProductName} chiếm {topProdShare}% doanh thu." : $"Giá trị đơn bình quân (AOV) đạt {avgOrderValue:N0} đ.")}";
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "INSIGHT 04 · BIÊN HIỆU QUẢ",
            Title = "Cơ Cấu Danh Mục & Tỷ Trọng Chi Phí",
            Content = effDesc,
            DetectedDate = today,
            Priority = hasCostData && grossMargin < 15 ? "Cao" : "Trung bình"
        });

        // INSIGHT 05 · HÀNH ĐỘNG 30 NGÀY
        string focusTarget = (worstRetStore != null && worstRetStore.ReturnRatePercent > 0) ? worstRetStore.StoreName
            : lowRevStore != null ? lowRevStore.StoreName
            : lowProd != null ? lowProd.ProductName
            : "chất lượng dữ liệu & hoàn trả";
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "INSIGHT 05 · HÀNH ĐỘNG 30 NGÀY",
            Title = $"Ưu Tiên Khắc Phục Ngắn Hạn: Tập trung {focusTarget}",
            Content = $"Kiểm tra nguyên nhân tại '{focusTarget}' và rà soát quy trình đối soát đơn hoàn/hủy để giữ tỷ lệ hoàn trả ({returnRate}%) dưới ngưỡng kiểm soát.",
            DetectedDate = today,
            Priority = "Cao"
        });

        // INSIGHT 06 · TẦM NHÌN 6–12 THÁNG
        string expandTarget = topProd != null ? $"nhóm {topProd.ProductName}"
            : bestStore != null ? $"đơn vị {bestStore.StoreName}"
            : topChan != null ? $"kênh {topChan.ChannelName}"
            : "quy trình dữ liệu chuẩn BI";
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "INSIGHT 06 · TẦM NHÌN 6–12 THÁNG",
            Title = $"Định Hướng Tối Ưu Trung & Dài Hạn: Mở rộng {expandTarget}",
            Content = $"Chuẩn hóa kiểm duyệt dữ liệu định kỳ trên Ngọc Linh Data Studio, tập trung nguồn lực mở rộng {expandTarget}{(bestStore != null ? $" và nhân rộng mô hình vận hành từ {bestStore.StoreName}" : "")}.",
            DetectedDate = today,
            Priority = "Trung bình"
        });

        var recommendations = structuredRecs.Select(r => $"[{r.Category.ToUpper()}] {r.Title}: {r.Content}").ToList();

        // Giai đoạn 4: Tìm nguyên nhân và Điểm bất thường (Insight Generation)
        var anomalies = new List<AnomalyInsight>();
        if (timeSeries.Count >= 3)
        {
            var avgDay = (double)timeSeries.Average(t => t.NetRevenue);
            var stdDay = Math.Sqrt(timeSeries.Average(t => Math.Pow((double)t.NetRevenue - avgDay, 2)));
            foreach (var pt in timeSeries)
            {
                if (stdDay > 0 && Math.Abs((double)pt.NetRevenue - avgDay) >= 1.15 * stdDay)
                {
                    bool isSpike = (double)pt.NetRevenue > avgDay;
                    var dayRows = valid.Where(x => x.OrderDate.Date.ToString("yyyy-MM-dd") == pt.Date).ToList();
                    var topDriverStore = dayRows.GroupBy(x => x.StoreName).OrderByDescending(g => Math.Abs(g.Sum(x => x.NetRevenue))).FirstOrDefault();
                    var topDriverProd = dayRows.GroupBy(x => x.ProductName).OrderByDescending(g => Math.Abs(g.Sum(x => x.NetRevenue))).FirstOrDefault();
                    anomalies.Add(new AnomalyInsight
                    {
                        Type = isSpike ? "Tăng vọt bất thường" : "Sụt giảm bất thường",
                        Dimension = $"Ngày {pt.Label}",
                        MetricValue = $"{pt.NetRevenue:N0} đ (TB ngày: {avgDay:N0} đ)",
                        RootCause = $"Nguyên nhân trực tiếp từ dữ liệu: chi nhánh '{topDriverStore?.Key}' ({topDriverStore?.Sum(x => x.NetRevenue):N0} đ) và sản phẩm '{topDriverProd?.Key}' ({topDriverProd?.Sum(x => x.NetRevenue):N0} đ).",
                        Severity = isSpike ? "Cơ hội" : "Cảnh báo"
                    });
                }
            }
        }
        foreach (var s in storeResults.Where(s => s.ReturnRatePercent > 10 || (hasCostData && s.MarginPercent < 12)))
        {
            anomalies.Add(new AnomalyInsight
            {
                Type = s.ReturnRatePercent > 10 ? "Hoàn trả cao bất thường" : "Biên lợi nhuận mỏng",
                Dimension = s.StoreName,
                MetricValue = $"Hoàn {s.ReturnRatePercent}% · Biên LN {(hasCostData ? $"{s.MarginPercent}%" : "N/A")}",
                RootCause = s.ReturnRatePercent > 10
                    ? $"Có {s.ReturnCount}/{s.OrderCount} đơn hoàn trả làm giảm doanh thu thực tế."
                    : $"Chiết khấu hoặc giá vốn cao kéo biên lợi nhuận xuống {s.MarginPercent}%.",
                Severity = "Cảnh báo"
            });
        }
        if (anomalies.Count == 0)
        {
            anomalies.Add(new AnomalyInsight
            {
                Type = "Chưa ghi nhận dị biệt lớn",
                Dimension = "Toàn hệ thống",
                MetricValue = $"{totalNet:N0} đ / {totalOrders} đơn",
                RootCause = timeSeries.Count < 3 ? "Dữ liệu chỉ có ít hơn 3 ngày nên chưa đủ chuỗi để phát hiện tăng/giảm đột biến theo ngày." : "Các điểm dữ liệu phân bổ trong ngưỡng ổn định.",
                Severity = "Ổn định"
            });
        }

        string execSummary = hasCostData
            ? $"Phân tích {valid.Count} bản ghi hợp lệ từ '{SourceName}': GMV đạt {totalGross:N0} đ, doanh thu thực tế {totalNet:N0} đ, lợi nhuận gộp {totalProfit:N0} đ (biên {grossMargin}%). Đã loại bỏ {LastStats.DuplicatesRemoved} đơn trùng, {LastStats.CancelledExcluded} đơn hủy/lỗi."
            : $"Phân tích {valid.Count} bản ghi hợp lệ từ '{SourceName}': GMV đạt {totalGross:N0} đ, doanh thu thực tế {totalNet:N0} đ (đã loại {LastStats.DuplicatesRemoved} đơn trùng, {LastStats.CancelledExcluded} đơn hủy/lỗi và đối soát {LastStats.ReturnsAdjusted} đơn hoàn). Chưa có dữ liệu giá vốn thực tế.";

        var completedOrders = valid.Count(x => x.Status == "COMPLETED");
        var totalUnitsSold = valid.Where(x => x.Status != "RETURNED").Sum(x => x.Quantity);
        var custGroups = schema.HasCustomer
            ? valid.Where(x => !string.IsNullOrWhiteSpace(x.CustomerName))
                   .GroupBy(x => x.CustomerName.Trim())
                   .Select(g => new CustomerAnalysisItem
                   {
                       CustomerName = g.Key,
                       OrderCount = g.Count(),
                       NetRevenue = g.Sum(x => x.NetRevenue),
                       TotalUnits = g.Sum(x => x.Quantity)
                   })
                   .OrderByDescending(c => c.NetRevenue)
                   .ToList()
            : new List<CustomerAnalysisItem>();

        return new DeepAnalysisResult
        {
            Schema = schema,
            CustomerPerformance = custGroups.Take(15).ToList(),
            TotalCustomers = custGroups.Count,
            RepeatCustomers = custGroups.Count(c => c.OrderCount > 1),
            CompletedOrders = completedOrders,
            TotalUnitsSold = totalUnitsSold,
            ExecutiveSummary = execSummary,
            OverallMetrics = new OverallMetrics
            {
                TotalGrossRevenue = totalGross,
                TotalDiscountAmount = totalDiscount,
                TotalNetRevenue = totalNet,
                TotalCost = totalCost,
                TotalGrossProfit = totalProfit,
                GrossMarginPercent = grossMargin,
                HasCostData = hasCostData,
                MissingCostCount = LastStats.ImputedCount,
                TotalOrders = totalOrders,
                ReturnOrders = returnOrders,
                ReturnRatePercent = returnRate,
                CancelRatePercent = cancelRate,
                PreventedOvercountValue = LastStats.PreventedOvercountValue,
                IsVerified = IsHumanVerified,
                IsEvaluated = IsResultEvaluated
            },
            PricingAnalysis = new PricingAnalysis
            {
                AverageOrderValue = avgOrderValue,
                AverageDiscountRate = avgDiscountRate,
                MaxDiscountPercent = maxDiscountOrderPct,
                DiscountedOrderCount = discountedOrderCount,
                TotalDiscountAmount = totalDiscount
            },
            StorePerformance = storeResults,
            ProductPerformance = productResults,
            ChannelBreakdown = channelResults,
            TimeSeries = timeSeries,
            GrowthRatePercent = growthRate,
            Anomalies = anomalies,
            StructuredRecommendations = structuredRecs,
            StrategicRecommendations = recommendations
        };
    }

    public object AnswerAnalystQuestion(string question)
    {
        var analysis = PerformDeepAnalysis();
        var q = question.ToLower();

        string answer;
        var matchedStore = analysis.StorePerformance.FirstOrDefault(s => q.Contains(s.StoreName.ToLower()));
        if (matchedStore != null)
        {
            answer = $"{matchedStore.StoreName}: doanh thu thuần {matchedStore.NetRevenue:N0} đ, lợi nhuận gộp {matchedStore.GrossProfit:N0} đ (biên {matchedStore.MarginPercent}%), {matchedStore.OrderCount} giao dịch, tỷ lệ hoàn {matchedStore.ReturnRatePercent}%. Đánh giá: {matchedStore.StatusAssessment}.";
        }
        else if (q.Contains("chiết khấu") || q.Contains("giá bán"))
        {
            answer = $"Giá trị đơn trung bình là {analysis.PricingAnalysis.AverageOrderValue:N0} đ. Tổng chiết khấu {analysis.PricingAnalysis.TotalDiscountAmount:N0} đ (tỷ lệ bình quân {analysis.PricingAnalysis.AverageDiscountRate}%, cao nhất {analysis.PricingAnalysis.MaxDiscountPercent}%).";
        }
        else if (q.Contains("tăng trưởng") || q.Contains("thời gian") || q.Contains("xu hướng"))
        {
            answer = analysis.TimeSeries.Count >= 2
                ? $"Qua {analysis.TimeSeries.Count} ngày ghi nhận, tốc độ biến động doanh thu từ ngày đầu đến ngày cuối là {analysis.GrowthRatePercent}%."
                : $"Dữ liệu hiện tập trung trong 1 ngày ({analysis.TimeSeries.FirstOrDefault()?.Label ?? "hôm nay"}) với doanh thu thuần {analysis.OverallMetrics.TotalNetRevenue:N0} đ.";
        }
        else if (q.Contains("doanh thu") || q.Contains("lợi nhuận") || q.Contains("kpi") || q.Contains("biên"))
        {
            answer = $"Tổng doanh thu thuần đạt {analysis.OverallMetrics.TotalNetRevenue:N0} đ, lợi nhuận gộp {analysis.OverallMetrics.TotalGrossProfit:N0} đ (biên lợi nhuận {analysis.OverallMetrics.GrossMarginPercent}%).";
        }
        else if (q.Contains("cửa hàng") || q.Contains("chi nhánh") || q.Contains("rủi ro"))
        {
            var bestStore = analysis.StorePerformance.FirstOrDefault();
            var riskStore = analysis.StorePerformance.OrderBy(s => s.MarginPercent).FirstOrDefault();
            answer = bestStore != null
                ? $"Chi nhánh dẫn đầu doanh thu là {bestStore.StoreName} ({bestStore.NetRevenue:N0} đ). Biên lợi nhuận thấp nhất là {riskStore?.StoreName} ({riskStore?.MarginPercent}%, hoàn {riskStore?.ReturnRatePercent}%)."
                : "Chưa có dữ liệu chi nhánh.";
        }
        else if (q.Contains("sản phẩm") || q.Contains("bán chạy") || q.Contains("sku"))
        {
            var topProd = analysis.ProductPerformance.FirstOrDefault();
            answer = topProd != null
                ? $"Sản phẩm dẫn đầu doanh thu là {topProd.ProductName} ({topProd.Sku}) với {topProd.QuantitySold} đơn vị, đạt {topProd.NetRevenue:N0} đ (biên {topProd.MarginPercent}%)."
                : "Chưa có dữ liệu sản phẩm.";
        }
        else if (q.Contains("hàng trả") || q.Contains("đổi trả") || q.Contains("hoàn"))
        {
            answer = $"Tỷ lệ hoàn trả hiện tại là {analysis.OverallMetrics.ReturnRatePercent}% ({analysis.OverallMetrics.ReturnOrders} đơn). Đã hạch toán giảm trừ doanh thu và hoàn nhập giá vốn.";
        }
        else
        {
            answer = $"{analysis.ExecutiveSummary} Bạn có thể hỏi chi tiết về doanh thu, lợi nhuận, chiết khấu, tăng trưởng, từng chi nhánh hoặc sản phẩm.";
        }

        return new
        {
            question,
            answer,
            keyMetrics = new
            {
                netRevenue = analysis.OverallMetrics.TotalNetRevenue,
                grossMargin = analysis.OverallMetrics.GrossMarginPercent,
                orders = analysis.OverallMetrics.TotalOrders
            }
        };
    }

    public string ExportToCsv()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Mã Đơn,Ngày Đơn,Khách Hàng,Chi Nhánh,SKU,Tên Sản Phẩm,Số Lượng,Đơn Giá,Chiết Khấu,Giá Vốn Đơn Vị,Doanh Thu Thuần,Tổng Giá Vốn,Lợi Nhuận Gộp,Biên LN %,Trạng Thái,Kênh,Ghi Chú");

        foreach (var r in CleanedData)
        {
            sb.AppendLine($"\"{r.OrderCode}\",\"{r.OrderDate:dd/MM/yyyy}\",\"{r.CustomerName}\",\"{r.StoreName}\",\"{r.Sku}\",\"{r.ProductName}\",{r.Quantity},{r.UnitPrice},{r.DiscountAmount},{r.UnitCost},{r.NetRevenue},{r.TotalCost},{r.GrossProfit},{r.MarginPercent},\"{r.Status}\",\"{r.Channel}\",\"{r.CleanRemark}\"");
        }

        return sb.ToString();
    }
}

public class CleaningStats
{
    public int RawCount { get; set; }
    public int CleanedCount { get; set; }
    public int DuplicatesRemoved { get; set; }
    public int CancelledExcluded { get; set; }
    public int ReturnsAdjusted { get; set; }
    public int MissingOrInvalid { get; set; }
    public int ImputedCount { get; set; }
    public decimal PreventedOvercountValue { get; set; }
}

public class DataSchemaProfile
{
    public bool HasDate { get; set; }
    public bool HasRevenue { get; set; }
    public bool HasCost { get; set; }
    public bool HasProduct { get; set; }
    public bool HasStore { get; set; }
    public bool HasCustomer { get; set; }
    public bool HasChannel { get; set; }
    public bool HasDiscount { get; set; }
    public bool HasStatus { get; set; }
    public List<string> DetectedColumns { get; set; } = new();
}

public class CustomerAnalysisItem
{
    public string CustomerName { get; set; } = "";
    public int OrderCount { get; set; }
    public decimal NetRevenue { get; set; }
    public int TotalUnits { get; set; }
}

public class DeepAnalysisResult
{
    public DataSchemaProfile Schema { get; set; } = new();
    public List<CustomerAnalysisItem> CustomerPerformance { get; set; } = new();
    public int TotalCustomers { get; set; }
    public int RepeatCustomers { get; set; }
    public int CompletedOrders { get; set; }
    public int TotalUnitsSold { get; set; }
    public string ExecutiveSummary { get; set; } = "";
    public OverallMetrics OverallMetrics { get; set; } = new();
    public PricingAnalysis PricingAnalysis { get; set; } = new();
    public List<StoreAnalysisItem> StorePerformance { get; set; } = new();
    public List<ProductAnalysisItem> ProductPerformance { get; set; } = new();
    public List<ChannelAnalysisItem> ChannelBreakdown { get; set; } = new();
    public List<TimeSeriesPoint> TimeSeries { get; set; } = new();
    public decimal GrowthRatePercent { get; set; }
    public List<AnomalyInsight> Anomalies { get; set; } = new();
    public List<StructuredRecommendation> StructuredRecommendations { get; set; } = new();
    public List<string> StrategicRecommendations { get; set; } = new();
}

public class AnomalyInsight
{
    public string Type { get; set; } = "";
    public string Dimension { get; set; } = "";
    public string MetricValue { get; set; } = "";
    public string RootCause { get; set; } = "";
    public string Severity { get; set; } = "";
}

public class OverallMetrics
{
    public decimal TotalGrossRevenue { get; set; }
    public decimal TotalDiscountAmount { get; set; }
    public decimal TotalNetRevenue { get; set; }
    public decimal TotalCost { get; set; }
    public decimal TotalGrossProfit { get; set; }
    public decimal GrossMarginPercent { get; set; }
    public bool HasCostData { get; set; }
    public int MissingCostCount { get; set; }
    public int TotalOrders { get; set; }
    public int ReturnOrders { get; set; }
    public decimal ReturnRatePercent { get; set; }
    public decimal CancelRatePercent { get; set; }
    public decimal PreventedOvercountValue { get; set; }
    public bool IsVerified { get; set; }
    public bool IsEvaluated { get; set; }
}

public class PricingAnalysis
{
    public decimal AverageOrderValue { get; set; }
    public decimal AverageDiscountRate { get; set; }
    public decimal MaxDiscountPercent { get; set; }
    public int DiscountedOrderCount { get; set; }
    public decimal TotalDiscountAmount { get; set; }
}

public class StoreAnalysisItem
{
    public string StoreName { get; set; } = "";
    public int OrderCount { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal GrossProfit { get; set; }
    public decimal MarginPercent { get; set; }
    public int ReturnCount { get; set; }
    public decimal ReturnRatePercent { get; set; }
    public string StatusAssessment { get; set; } = "";
}

public class ProductAnalysisItem
{
    public string Sku { get; set; } = "";
    public string ProductName { get; set; } = "";
    public int QuantitySold { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal GrossProfit { get; set; }
    public decimal DiscountAmount { get; set; }
    public decimal MarginPercent { get; set; }
}

public class ChannelAnalysisItem
{
    public string ChannelName { get; set; } = "";
    public int OrderCount { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal RevenueSharePercent { get; set; }
}

public class TimeSeriesPoint
{
    public string Date { get; set; } = "";
    public string Label { get; set; } = "";
    public decimal NetRevenue { get; set; }
    public decimal GrossProfit { get; set; }
    public int Orders { get; set; }
}

public class StructuredRecommendation
{
    public string Category { get; set; } = "";
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
    public string DetectedDate { get; set; } = "";
    public string Priority { get; set; } = "Trung bình";
}
