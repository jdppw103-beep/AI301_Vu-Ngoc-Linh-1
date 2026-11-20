# DataLens AI — Hệ thống phân tích dữ liệu khách hàng

Giao diện web sạch, khởi động **không có dữ liệu sẵn**, chờ khách hàng tải tệp lên và xử lý qua 4 bước:

**Kết nối → Làm sạch → Phân tích → Xuất kết quả**

## Khởi chạy
- Nhấp đúp `start.bat`, hoặc: `dotnet run --urls "http://localhost:5000"`
- Mở trình duyệt (Chrome/Edge): **http://datalens.localhost:5000**
  (vẫn truy cập được qua http://localhost:5000)

## Định dạng tệp đầu vào
CSV (`,` `;` hoặc tab) hoặc JSON. Bấm **“Tệp template”** trên giao diện để tải tệp mẫu chỉ có tiêu đề cột:

```
OrderCode,OrderDate,CustomerName,StoreName,Sku,ProductName,Quantity,UnitPrice,DiscountAmount,UnitCost,Status,Channel
```

Hệ thống tự nhận diện cả tên cột tiếng Việt (Mã đơn, Chi nhánh, Số lượng, Đơn giá, Giá vốn, Trạng thái, Kênh…).
Trạng thái được chuẩn hoá: `hủy/cancel` → CANCELLED, `lỗi/fail` → FAILED, `hoàn/trả/return` → RETURNED.

## Cấu trúc
- `Program.cs` — backend ASP.NET Core: đọc tệp, làm sạch, phân tích, xuất
- `wwwroot/index.html` — giao diện
- `DataLensAI.csproj` — dự án .NET 8
