# HƯỚNG DẪN CÀI ĐẶT & SỬ DỤNG HỆ THỐNG WEB ÔN THI THPT MÔN HÓA HỌC 2026-2027
**TRƯỜNG THPT QUANG TRUNG**

---

## 1. GIỚI THIỆU TỔNG QUAN
Hệ thống WebApp ôn thi THPT môn Hóa học được xây dựng hoàn chỉnh dưới dạng Single-Page Application (HTML/CSS/JavaScript thuần), tích hợp đầy đủ:
- **3 đề thi chuẩn hóa (ĐỀ 01, ĐỀ 02, ĐỀ 03)** theo đúng ma trận và định dạng đề thi Tốt nghiệp THPT mới nhất của Bộ GD&ĐT (gồm 3 phần: Phần I - 18 câu nhiều lựa chọn, Phần II - 4 câu đúng/sai, Phần III - 6 câu trả lời ngắn).
- **Hệ thống dành cho Học sinh**: Đăng nhập (Họ tên, Lớp), chọn đề thi, đồng hồ đếm ngược 50 phút, thanh palette chuyển câu, nộp bài, xem điểm tức thời và xem hướng dẫn giải chi tiết từng câu.
- **Hệ thống dành cho Giáo viên (Mật khẩu: `giaovien123`)**: Thống kê phổ điểm, bảng điểm chi tiết của từng học sinh, xem chi tiết từng câu làm đúng/sai, xuất file Excel (`.csv` chuẩn UTF-8 tiếng Việt), và nút **Xóa dữ liệu cũ để làm lại từ đầu với lớp khác**.
- **Đồng bộ Google Sheets thời gian thực**: Tự động lưu kết quả bài làm vào Google Sheet ngay khi học sinh bấm nộp bài, đồng bộ dữ liệu hai chiều và xóa dữ liệu trực tiếp trên Google Sheet.

---

## 2. CÁCH MỞ VÀ SỬ DỤNG ỨNG DỤNG
1. Mở tệp `index.html` hoặc `on_thi_thpt_hoa_hoc_2026_quang_trung.html` bằng bất kỳ trình duyệt web nào (Google Chrome, Microsoft Edge, Firefox, Cốc Cốc, Safari trên máy tính, máy tính bảng hoặc điện thoại).
2. Trang web chạy 100% độc lập, không yêu cầu cài đặt phần mềm máy chủ hay môi trường phức tạp.

---

## 3. HƯỚNG DẪN KẾT NỐI GOOGLE SHEETS ĐỒNG BỘ TỨC THỜI (DÀNH CHO GIÁO VIÊN)

Để kết quả bài làm của học sinh tự động lưu trực tiếp vào Google Sheets của Thầy/Cô:

### Bước 1: Tạo Google Sheet mới
1. Đăng nhập Google Drive và tạo một bảng tính **Google Sheets** mới.
2. Đặt tên trang tính tùy ý (ví dụ: *Kết quả Thi Hóa THPT Quang Trung 2026-2027*).

### Bước 2: Dán mã Google Apps Script
1. Trên Google Sheets, nhấn menu **Tiện ích mở rộng (Extensions)** > **Apps Script**.
2. Xóa toàn bộ nội dung mẫu có sẵn, sao chép và dán toàn bộ đoạn mã bên dưới vào:

```javascript
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var contents = e.postData.contents;
    var data = JSON.parse(contents);

    // Xóa toàn bộ dữ liệu khi giáo viên bấm "Xóa dữ liệu cũ" trên Web
    if (data.action === "clear_all") {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Đã làm sạch dữ liệu cũ trên Google Sheets."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Tự động tạo dòng tiêu đề nếu bảng tính còn trống
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Thời điểm nộp",
        "Họ và tên thí sinh",
        "Lớp",
        "Đề thi",
        "Tổng điểm (Thang 10)",
        "Điểm Phần I (TN)",
        "Điểm Phần II (Đ/S)",
        "Điểm Phần III (Trả lời ngắn)",
        "Thời gian làm bài",
        "Ghi chú"
      ]);
      var headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#0d274c");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // Ghi nhận bài làm mới của học sinh
    var d = data.data;
    var newRow = [
      d.timestamp || new Date().toLocaleString("vi-VN"),
      d.student_name,
      d.student_class,
      d.exam_id,
      d.total_score,
      d.p1_score,
      d.p2_score,
      d.p3_score,
      d.duration,
      "Tự động gửi từ WebApp"
    ];
    sheet.appendRow(newRow);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Đã lưu kết quả thành công."
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    var results = [];

    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (r[1] && r[1].toString().trim() !== "") {
          results.push({
            timestamp: r[0],
            student_name: r[1],
            student_class: r[2],
            exam_id: r[3],
            total_score: Number(r[4]) || 0,
            p1_score: Number(r[5]) || 0,
            p2_score: Number(r[6]) || 0,
            p3_score: Number(r[7]) || 0,
            duration: r[8]
          });
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      count: results.length,
      data: results
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
```

### Bước 3: Triển khai Web App (Deploy)
1. Bấm nút **Triển khai (Deploy)** ở góc trên bên phải > chọn **Triển khai mới (New deployment)**.
2. Chọn loại: **Ứng dụng web (Web app)**.
3. Thiết lập thông số:
   - **Mô tả**: *Ôn thi Hóa Quang Trung*
   - **Thực thi dưới dạng (Execute as)**: *Tôi (địa chỉ email của bạn)*
   - **Ai có quyền truy cập (Who has access)**: **Bất kỳ ai (Anyone)** *(Rất quan trọng để học sinh có thể gửi bài mà không cần đăng nhập Google)*.
4. Bấm **Triển khai (Deploy)**. Nếu Google yêu cầu quyền truy cập, chọn tài khoản và cấp quyền.
5. Sao chép chuỗi **URL ứng dụng web (Web App URL)** có dạng `https://script.google.com/macros/s/.../exec`.

### Bước 4: Dán URL vào WebApp
1. Mở trang web ôn thi, bấm **Dành cho Giáo viên** ở góc phải trên.
2. Nhập mật khẩu: `giaovien123`.
3. Kéo xuống phần **Kết Nối Google Sheets Tức Thời (Live Sync)**.
4. Dán URL vừa sao chép vào ô và nhấn **Lưu cấu hình**.
5. Kể từ thời điểm này, mỗi khi bất kỳ học sinh nào nhấn **Nộp bài**, điểm số và thông tin chi tiết sẽ được tự động ghi vào dòng mới trên Google Sheets trong chớp mắt!

---

## 4. QUẢN LÝ DỮ LIỆU & LÀM LẠI TỪ ĐẦU VỚI LỚP KHÁC
- Khi chuyển sang cho lớp khác thi (ví dụ từ 12A1 sang 12A2), giáo viên chỉ cần bấm nút **"Xóa dữ liệu cũ (Làm lại với lớp khác)"** trong trang Giáo viên.
- Hệ thống sẽ xóa sạch dữ liệu học sinh lớp cũ cả trên trình duyệt lẫn trên Google Sheets, sẵn sàng cho lớp mới thi bài.
