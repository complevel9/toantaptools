## Vài chương trình xử lý văn bản

### Các phần mềm cần cài

- Hệ điều hành Windows
- MS Word 2010 (các bản mới hơn chắc là cũng dùng được)
- [Python](https://www.python.org/downloads/) + `pip install PyMuPDF python-docx pywin32`
- [ghostscript](https://ghostscript.com/releases/gsdnld.html) (Không cần cài nếu không dùng thính năng gắn tài liệu của `pdfsplit.py`)


### `pdfsplit.py`: cắt mỗi trang PDF hai cột thành hai trang PDF một cột

Lệnh: `pdfsplit.py input.pdf -o output.pdf [-ra] [-rb]`

Chức năng: Đọc file `input.pdf` là tệp tài liệu hai cột, cắt giữa mỗi trang thành hai trang, xuất ra thành `output.pdf`. Có thể cho thêm `-ra` vào lệnh để bỏ qua nửa trang đầu, thêm `-rb` để bỏ qua nửa trang cuối.

### `pdfsplit.py`: chọn một số trang của tài liệu PDF xong gắn vào cuối của tài liệu khác

Lệnh: `pdfsplit.py input.pdf -o output.pdf -rr left_page_num right_page_num appendto.pdf [--nogs]`

Chức năng: Đọc từ trang `left_page_num` đến trang `right_page_num` của tệp `input.pdf`, rồi gắn vào đuôi của tệp `appendto.pdf`, rồi dùng ghostscript để dọn phần phông chữ lặp trong quá trình gắn 2 tệp vào nhau, rồi xuất kết quả ra `output.pdf`. Nếu muốn bỏ qua bước dọn dẹp bằng ghostscript thì có thể thêm `--nogs` vào lệnh.

### `markers.py`: in ra tiêu đề, số chú thích trong văn bản
Lệnh: `markers.py file.doc` hoặc là `markers.py docfolder`

Chức năng: In ra các số chú thích (endnote markers) và các tiêu đề tìm được trong một văn bản `file.doc` hoặc tất cả các văn bản `.doc` thuộc thư mục `docfolder`. Script này tìm số chú thích bằng cách phát hiện các số superscript, tìm tiêu đề là một định dạng hardcode trong script. Mục đích chính là để tóm tắt xem văn bản nào có những đề mục gì bên trong với endnote từ bao nhiêu đến bao nhiêu.

### `toantap-macros.dotm`: add-in trong Word để đánh số lại số trang, số endnote

Cách cài đặt: copy tệp này vào thư mục `%APPDATA%\Microsoft\Word\STARTUP\`, nếu không có thì tự tạo thư mục.