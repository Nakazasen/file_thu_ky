# Smart Factory — Công cụ quản lý đa ngữ (Nhật/Trung/Việt)

Tài liệu tổng hợp toàn bộ việc đã làm: đọc hiểu file gốc → triển khai đa ngữ cho nội dung mới → backup chống xóa → điều tra vệt gạch.

## 1. File trong thư mục

| File | Trạng thái | Ghi chú |
|---|---|---|
| `スマートファクトリー分科会_管理ツール案 のコピー.xlsx` | Gốc (giữ nguyên) | Bản sếp đưa, 10 sheets, 136 KB |
| `backup_goc_20260929.xlsx` | Backup gốc | Copy 1:1 trước khi sửa |
| `smart_factory_3ngu_v3.xlsx` | **File chính đang dùng** | v2 + nới dòng ví dụ (48/60pt). 11 sheets, đầy đủ nhất |
| `smart_factory_3ngu_v2.xlsx` | Dự phòng | Nội dung như v3, dòng ví dụ còn thấp (31.5pt) |
| `SmartFactory_VN.dic` | Từ điển team (409 từ) | Nạp 1 lần/máy để hết gạch đỏ. Xem mục 5 |
| `BROKEN_do_not_use_*.xlsx` | HỎNG — đừng mở | File thử nghiệm bị lỗi, giữ lại để đối chứng |
| `README.md` | Tài liệu này | |

## 2. Nguyên lý chuyển ngữ của file gốc (đã đọc & xác minh)

- Sheet `言語切替!B3` là dropdown `日本語,中文,Tiếng Việt` — nút chuyển duy nhất.
- Sheet `_翻訳マスタ` là từ điển: cột A = KEY (mã), B/C/D = Nhật/Trung/Việt (~247 key: TITLE, PURPOSE, DASH_*, M_*, WBS_*, ISSUE_*, DEC_*, KPI_*, REV_*, ALERT_*...).
- Mọi tiêu đề/header ở các sheet đều là công thức, ví dụ:
  `=INDEX(_翻訳マスタ!$B$2:$D$32,MATCH("TITLE",_翻訳マスタ!$A$2:$A$32,0),MATCH(言語切替!$B$3,_翻訳マスタ!$B$1:$D$1,0))`
  → đổi B3 là toàn bộ tiêu đề nhảy theo ngay (đã kiểm chứng hiển thị tiếng Việt khi B3 = Tiếng Việt).
- Sheet `02_マスタ` (danh mục WG/đơn vị/trạng thái...) bản thân cũng dịch qua INDEX, nên các dropdown nhập liệu (Data Validation dạng x14 trỏ về `02_マスタ`) cũng đổi theo.
- Các cột đếm/tính (`01` COUNTIFS, `03!N`判定, `04!S`アラート) so sánh với INDEX theo ngôn ngữ hiện tại chứ không so chuỗi cứng.
- **Giới hạn gốc:** chỉ label dịch được. Nội dung người dùng gõ tay (tên task, tiêu đề vấn đề...) là text tự do → nút chuyển không tác dụng. Tên sheet không đổi.

## 3. Yêu cầu của sếp & phương án đã chốt

1. Nội dung mới gõ vào cũng phải dịch được (Copilot/TRANSLATE).
2. Song song 3 ngôn ngữ, kể cả nội dung mới.
3. Chống xóa nhầm → backup lịch sử.

Phương án (giữ nguyên cột nhập, thêm cột dịch song song — ít rủi ro nhất):
mỗi trường tự do giữ nguyên cột nhập (người dùng gõ như cũ) + 3 cột phụ JP/CN/VI dùng hàm `TRANSLATE`:
`=IF(E5="","",IFERROR(TRANSLATE(E5,,"ja"),""))`
(có `IF` chặn ô trống để khỏi tốn quota, `IFERROR` để Excel cũ hiện trống thay vì báo lỗi; cần M365 + mạng, dưới 3000 ký tự/ô).

## 4. Đã triển khai vào file (v2/v3)

- **03_WBS:** `E Task → Q/R/S`, `P Note → T/U/V` (6 cột, tới dòng 204).
- **04_課題ToDo:** `E/F/G/O/P → U–AI` (15 cột, tới dòng 304).
- **05_決定事項:** `D/E/F/J/O → P–AD` (15 cột, tới 204).
- **06_KPI:** `D/E/F/Q/R → S–AG` (15 cột, tới 204).
- **07_更新履歴:** `E/F/H → I–Q` (9 cột, tới 104).
- Header các cột phụ tự đổi theo nút ngôn ngữ (tên trường + hậu tố JP/CN/VI).
- `_翻訳マスタ` thêm 15 key mới (`PARA_*`, `BACKUP_*`, `GOMI_*`); `言語切替` thêm bảng mã `ja / zh-Hans / vi` + ghi chú M365.
- `00_運用ルール` thêm mục 5 (dịch song song) + mục 6 (chống xóa), mỗi sheet nhập thêm dòng hướng dẫn ở A3.
- Sheet mới `_ゴミ箱_Backup` (Ngày xóa/Trang/Mã/Snapshot/Người xóa/Lý do) + quy tắc: cấm xóa dòng vật lý, chuyển `Hoàn thành/Rút lại` rồi copy sang đây; file đưa lên SharePoint bật Version History 500 bản, hàng tuần lưu bản riêng.
- Đã test thật: gõ tiếng Việt vào `04!E9` → `U9` ra tiếng Nhật, `V9` ra tiếng Trung ngay.
- Validation dropdown gốc giữ nguyên 100% (kiểm bằng đọc trực tiếp XML).

> Sự cố giữa chừng: 1 lần tạo sheet bằng COM sai cách làm `_翻訳マスタ` bị đổi tên nhầm → công thức `#N/A` hàng loạt. Đã khôi phục từ backup và làm lại sang file tên ASCII (`smart_factory_3ngu_v2/v3`) bằng cách an toàn. File hỏng giữ lại tên `BROKEN_*` để tránh mở nhầm.

## 5. Vụ "vệt gạch ngang" — kết luận điều tra

Hiện tượng: vệt ngang đè lên chữ ở ô xanh (ô công thức), ô mã/số gõ tay thì không; mở lại là thấy; nhìn "giống lỗi công thức".

Đã loại trừ (có bằng chứng): strikethrough/underline/CF/shape/hyperlink/table/rich-text trong file = 0; ký tự tổ hợp ẩn = không có; font/viền sạch; add-in đã test bật/tắt.

Thí nghiệm quyết định: cùng 1 ô, chữ Anh thì sạch, chữ Việt thì vệt → **vệt phụ thuộc nội dung = do bộ kiểm tra chính tả/ngữ pháp (proofing) của máy**, không phải lỗi file:
- Máy chỉ có bộ chính tả Anh/Pháp/TBN, không có bộ Việt (`CheckSpelling` chữ Việt = False) → chữ Việt nào cũng bị đánh dấu.
- Phân tích pixel ảnh chụp 24-bit: 3400 pixel XANH, 0 pixel đỏ → vệt là gạch xanh ngữ pháp của Microsoft Editor.
- Đã nới dòng ví dụ lên 48/60pt (v3) nên vệt nằm dưới chân chữ, không còn cắt ngang chữ.

Đã làm để giảm: nạp 409 từ Việt cố định của file vào từ điển sống `RoamingCustom.dic` (có backup `.bak_20260929`), xuất `SmartFactory_VN.dic` cho team (máy khác: mở Notepad copy vào RoamingCustom.dic của máy đó, hoặc `File > Options > Proofing > Custom Dictionaries > Add`).

Hết hẳn vệt xanh trên từng máy (30 giây, không cần admin): mở Excel → `Review > Editor` → cuối ngăn bấm `Settings` (bánh răng) → bỏ tick `Grammar`. Triệt để lâu dài: cài Vietnamese proofing (`File > Options > Language`). Vệt kiểm tra **không bao giờ in ra** nên demo với sếp cứ dùng Print Preview/PDF.

## 6. Cách dùng & chuyển giao team

1. Dùng file `smart_factory_3ngu_v3.xlsx`. Nhập liệu như cũ vào cột xanh; đọc bản dịch ở cụm JP/CN/VI bên phải (kéo thanh trượt sang phải).
2. Nút `言語切替!B3` đổi tiêu đề/header toàn file.
3. Không xóa dòng vật lý — chuyển trạng thái + copy sang `_ゴミ箱_Backup`.
4. Mỗi máy team: nạp `SmartFactory_VN.dic` 1 lần + tắt Grammar trong Editor Settings nếu còn vệt.
5. Đưa file lên SharePoint/Teams để có Version History; mọi sửa lớn trong file phải qua COM cẩn thận (đã từng lỗi đổi tên sheet) hoặc báo trước.
