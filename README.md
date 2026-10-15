# Korean Grammar Extractor

Ứng dụng Flask tìm cấu trúc phủ định kép trong JSON hoặc ZIP chứa JSON,
hiển thị kết quả có phân trang và tải Excel.

## Chạy trên Render với RAM 512 MB

Trong **Settings → Build & Deploy**, đặt:

- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `gunicorn app:app`
- Giữ **1 instance**. Nếu đã đặt biến `WEB_CONCURRENCY` hoặc `GUNICORN_CMD_ARGS`
  để tăng số worker, hãy bỏ cấu hình tăng worker đó.

Commit và đẩy các file đã sửa lên repository Render đang dùng, rồi deploy lại.
Render dùng Start Command trong dashboard; sửa riêng Procfile sẽ không ghi đè
Start Command cũ như `gunicorn app:app --workers 2`.

`gunicorn.conf.py` tự được đọc từ thư mục dự án: 1 worker, 4 thread,
timeout 600 giây, bind theo biến `PORT`. Một worker cũng đảm bảo các request
upload, xem trạng thái và tải kết quả truy cập cùng bộ quản lý tác vụ.

## Những thay đổi giảm RAM

- Upload được ghi thẳng xuống file tạm, kể cả khi tải nhiều JSON nhỏ.
- ZIP được mở từ đĩa; từng JSON được sao chép theo khối 64 KiB vào file tạm.
  Không giữ toàn bộ dữ liệu sau giải nén trong RAM và không giải nén theo đường
  dẫn lấy từ tên thành viên ZIP.
- `ijson` đọc JSON theo luồng, chỉ lấy các trường cần thiết của từng phát ngôn.
  Hỗ trợ `document` và `utterance` dạng object hoặc array như bản cũ.
  JSON được đọc hai lượt để lấy đúng metadata kể cả khi metadata nằm cuối file.
- Kết quả lưu trong SQLite trên đĩa, cache khoảng 2 MiB mỗi kết nối. Phân trang
  và lọc bằng truy vấn SQL; API giới hạn tối đa 200 kết quả mỗi trang.
- Excel được ghi từng dòng bằng `openpyxl` với `write_only=True`, ra file tạm
  trên đĩa. Không tạo workbook chứa toàn bộ ô hoặc bản Excel trong RAM.
- Chỉ nhận một tác vụ xử lý và một lượt xuất Excel cùng lúc; yêu cầu tiếp theo
  được báo bận. JSON lỗi được bỏ qua và đếm; lỗi tác vụ được hiển thị trên web.
- Xóa upload khi xử lý xong. Kết quả được giữ 1 giờ, rồi được dọn khi có request
  xử lý/xem trạng thái/xem kết quả/tải xuống tiếp theo.

Không thể cam kết mọi dữ liệu đều dưới 512 MB: một trường chuỗi cực lớn hoặc
ZIP có số lượng thành viên cực lớn vẫn cần RAM. Bản sửa đổi phần bộ nhớ tăng
theo toàn bộ corpus/kết quả sang xử lý tuần tự. Đổi lại cần dung lượng đĩa tạm
cho ZIP, JSON đang xử lý, SQLite và Excel; CPU phải đọc JSON hai lượt.

Đường dẫn như `E:\data` chỉ dùng khi chạy ứng dụng trên chính máy chứa thư mục.
Khi dùng Render, hãy upload ZIP/JSON; Render không đọc được ổ đĩa máy cá nhân.
File tạm và kết quả không được đảm bảo giữ lại sau restart/redeploy; tải Excel
trước khi hết thời gian giữ kết quả. Excel có tối đa 1.048.575 dòng dữ liệu mỗi
file; với kết quả lớn hơn hãy tải theo từng cấu trúc.

## Chạy và kiểm thử trên máy

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
```

Mở `http://localhost:5000`. Kiểm thử luồng ZIP, JSON, bộ lọc, phân trang,
Excel, lỗi và giới hạn tác vụ:

```powershell
.venv\Scripts\python.exe -m unittest -v test_low_memory
```

Đo RAM của tiến trình Python với ZIP giả lập chứa hơn 512 MiB JSON và xuất
100.000 dòng Excel (cần khoảng 1 GB đĩa tạm; không dùng dữ liệu cá nhân):

```powershell
.venv\Scripts\python.exe -m pip install psutil
.venv\Scripts\python.exe benchmark_memory.py --rows 100000 --padding 6000
```

Benchmark báo RSS cao nhất lấy mẫu mỗi 10 ms, kích thước JSON/ZIP/Excel và
thời gian chạy. Đây là kiểm tra trên máy hiện tại, chưa bao gồm Gunicorn master,
Render/Linux và dữ liệu thực tế. Sau deploy, kiểm tra lại mục Metrics của Render
trong lúc upload, xử lý và tải Excel.

Kết quả kiểm tra ngày 16/10/2026 trên Windows, Python 3.14.3: JSON 579,55 MiB,
ZIP 0,78 MiB, Excel 100.000 dòng / 3,18 MiB, RSS đỉnh tiến trình Python
55,74 MiB, thời gian 92,42 giây. Dữ liệu giả lập có trường phụ lặp lại nên
ZIP nén rất mạnh; số liệu này kiểm chứng luồng đọc và xuất, không đại diện cho
mọi corpus và không phải số đo tổng bộ nhớ dịch vụ Render.

## Tài liệu tham khảo

- [Render: deploy Flask](https://render.com/docs/deploy-flask)
- [Gunicorn: threads giảm bộ nhớ so với nhiều process](https://docs.gunicorn.org/en/stable/design.html#how-many-threads)
- [openpyxl: write-only mode](https://openpyxl.readthedocs.io/en/stable/optimized.html)
- [ijson: streaming JSON](https://pypi.org/project/ijson/)
