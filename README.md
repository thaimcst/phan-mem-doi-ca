# 🔄 Quản Lý Đổi Ca

Ứng dụng web quản lý đơn xin đổi ca làm việc cho Đội Vận hành, bảo trì đường hầm.

## Tính năng

- ✅ Tạo đơn xin đổi ca (Nhân viên / Trưởng ca)
- 💾 Lưu đơn lên Firebase Realtime Database
- 📋 Xem lịch sử tất cả đơn đã tạo
- 📄 Xuất file Word (.docx) cho từng đơn
- 🖨️ In đơn trực tiếp từ trình duyệt
- 📱 Responsive – hỗ trợ điện thoại và máy tính bảng

## Cấu trúc thư mục

```
quan-ly-doi-ca/
├── index.html   # Giao diện chính
├── style.css    # Toàn bộ CSS / giao diện
├── app.js       # Logic xử lý, kết nối Firebase, xuất Word
└── README.md
```

## Cài đặt & Chạy

Không cần cài đặt. Chỉ cần mở `index.html` trong trình duyệt, hoặc deploy lên **GitHub Pages**:

1. Push toàn bộ thư mục lên GitHub repo
2. Vào **Settings → Pages**
3. Chọn branch `main`, thư mục `/ (root)`
4. Lưu → GitHub cấp link truy cập công khai

## Cấu hình Firebase

Mở `app.js` và thay đường dẫn API nếu cần:

```js
const API = 'https://<your-project>.firebasedatabase.app/don_doi_ca';
```

## Công nghệ sử dụng

- HTML5 / CSS3 / Vanilla JavaScript
- [Firebase Realtime Database](https://firebase.google.com/)
- [JSZip](https://stuk.github.io/jszip/) – xuất file .docx
- [Google Fonts – Inter](https://fonts.google.com/specimen/Inter)
