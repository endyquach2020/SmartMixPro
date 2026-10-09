# SmartMix Pro

Ứng dụng web hỗ trợ trộn đề thi trắc nghiệm từ file Word (.docx) và xuất ma trận đáp án Excel (.xlsx) chuẩn Bộ Giáo dục & Đào tạo.

## Tính năng chính
- **Trộn đề thông minh theo nhóm mã**:
  - `<#g1>`: Trắc nghiệm 4 lựa chọn (hoán vị câu hỏi và phương án A, B, C, D).
  - `<#g2>`: Trắc nghiệm đúng / sai (hoán vị các ý con a, b, c, d).
  - `<#g3>`: Câu hỏi trả lời ngắn (hoán vị câu hỏi, trích xuất đáp án vào Excel).
  - `<#g4>`: Nhóm câu hỏi cố định (tự luận, bài đọc hiểu, bài nghe...).
- **Xuất file chuẩn**:
  - Đề gốc mã 000 và các mã đề hoán vị định dạng Word (.docx).
  - Bảng ma trận đáp án định dạng Excel (.xlsx).
  - Tải trọn bộ đề thi và đáp án dạng nén (.zip).
- **Giao diện hiện đại & tiện ích**: Chạy trực tiếp trên trình duyệt, không cần cài đặt phần mềm phức tạp.

## Hướng dẫn chạy cục bộ (Local)
1. Nhấp đúp vào file `chay-web.bat` hoặc chạy lệnh:
   ```bash
   python -m http.server 8000
   ```
2. Mở trình duyệt và truy cập: `http://localhost:8000`

## Tác giả
- Email: `nhicnttcantho@gmail.com`
- Zalo: `0917809488`
