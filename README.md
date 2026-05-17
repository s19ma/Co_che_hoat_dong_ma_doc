# BÁO CÁO PHÂN TÍCH DATASET VÀ CÔNG CỤ TRÍCH XUẤT MACRO MALWARE

## Mục lục
1. [Tổng quan Dataset](#1-tổng-quan-dataset)
2. [Công cụ phân tích và trích xuất (Tools)](#2-công-cụ-phân-tích-và-trích-xuất-tools)
3. [Dữ liệu mẫu lành tính (Benign VBA)](#3-dữ-liệu-mẫu-lành-tính-benign-vba)
4. [Dữ liệu cập nhật từ MalwareBazaar Daily](#4-dữ-liệu-cập-nhật-từ-malwarebazaar-daily)

---

## 1. Tổng quan Dataset

*   **Nguồn dữ liệu gốc:** [GitHub - Co_che_hoat_dong_ma_doc](https://github.com/s19ma/Co_che_hoat_dong_ma_doc/tree/main/All_Excel_Doc_Macro)

### 1.1. Phân loại các nhóm mẫu độc hại (Malware Samples)

| Nhóm | Số lượng | Kỹ thuật cốt lõi / Đặc điểm nhận dạng |
| :--- | :---: | :--- |
| **Nhóm 1** | 4 mẫu | Sử dụng hàm `EXEC` để thực thi các lệnh shell (ví dụ: khởi chạy `powershell.exe` với lệnh mã hóa). |
| **Nhóm 2** | 4 mẫu | Sử dụng hàm `REGISTER` để gọi các hàm API từ thư viện hệ thống `Kernel32.dll` (`VirtualAlloc`, `WriteProcessMemory`, `CreateThread`). Cơ chế này ghi mã độc trực tiếp vào bộ nhớ và thực thi (thường gây treo Excel). |
| **Nhóm 3** | 5 mẫu | Sử dụng hàm `CALL` để tương tác với thư viện `urlmon.dll` nhằm tải tệp độc hại từ bên ngoài (ví dụ: từ GitHub) và kích hoạt thông qua `shell32.dll`. |
| **Nhóm 4** | 4 mẫu | Kết hợp các kỹ thuật của Nhóm 1, 2, 3 nhưng được tích hợp thêm **cơ chế phát hiện môi trường phân tích** (Anti-Sandbox / Anti-VM). |
| **Nhóm 5** | 16 mẫu | Các mẫu mã độc VBA/Doc/Excel ở trạng thái *Chưa phân loại* rõ hành vi chính. |
| **Nhóm 6** | 78 mẫu | Các mẫu VBA thực tế được thu thập trực tiếp từ [MalwareBazaar](https://bazaar.abuse.ch/browse/tag/vba/) (Chưa phân loại sẵn). |

> **Đặc điểm chung của Nhóm 6:**
> * Tần suất sử dụng kỹ thuật làm rối mã (obfuscation) rất cao.
> * Kết hợp linh hoạt nhiều hành vi nguy hiểm cùng lúc: tải payload, thực thi lệnh shell, inject memory.
> * Trang bị các kỹ thuật né tránh sandbox (anti-analysis).

### 1.2. Các kỹ thuật lẩn trốn (Evasion Techniques) được phát hiện

*   **Làm rối mã (Obfuscation):**
    *   Sử dụng các hàm có sẵn của Excel như `CHAR`, `MID`, `CODE`, `HEX2DEC` để băm nhỏ và che giấu chuỗi lệnh độc hại gốc.
    *   Dùng hàm `FORMULA` để dựng động nội dung mã độc ngay trong thời gian thực (runtime) khi người dùng vừa mở file.
    *   Sử dụng các vòng lặp `While` phức tạp nhằm tạo và gọi các hàm thực thi một cách gián tiếp.
*   **Ẩn giấu thành phần (Hidden Sheets):**
    *   Giấu mã macro trong các bảng tính (Sheets) được thiết lập thuộc tính ẩn `Hidden` hoặc siêu ẩn `VeryHidden`.
    *   *Lưu ý từ tài liệu:* Việc lạm dụng thuộc tính `VeryHidden` thường làm tăng đáng kể điểm nghi vấn (score) và dễ bị các phần mềm diệt virus (AV) hiện đại gắn cờ cảnh báo.
*   **Phát hiện môi trường ảo/Sandbox (Sandboxing Detection):**
    *   Sử dụng các hàm hệ thống như `GET.WORKSPACE`, `GET.DOCUMENT`, `GET.WINDOW` để thu thập thông tin cấu hình phần cứng và OS.
    *   **Kiểm tra điều kiện môi trường thực:** Mã độc chỉ kích hoạt nếu thỏa mãn các yếu tố thực tế như: có hành vi di chuyển chuột, cửa sổ Excel được phóng to tối đa (`Maximised`), hoặc kiểm tra thiết bị có khả năng ghi âm/phát thanh hay không.

---

### Hình ảnh minh họa cấu trúc dữ liệu
<img width="1536" height="2752" alt="Cấu trúc thư mục dataset" src="https://github.com/user-attachments/assets/b75e7cff-4593-4270-97ab-b02274941f7e" />
<img width="1553" height="645" alt="Tổng quan danh sách mẫu độc hại" src="https://github.com/user-attachments/assets/ca86ebb0-3de9-4489-9980-c7500af580e9" />

---

## 2. Công cụ phân tích và trích xuất (Tools)

### 2.1. Công cụ phân tích tĩnh: `olevba`
*   **Nguồn:** [oletools - olevba](https://github.com/decalage2/oletools/wiki/olevba)
*   **Cách thức hoạt động:**
    1.  Tự động quét và trích xuất nguyên bản (raw) các đoạn mã macro VBA từ tệp tin Microsoft Office.
    2.  Phân tích cú pháp và gắn cờ cảnh báo các thành phần:
        *   **Chuỗi đáng ngờ (Suspicious strings):** Các đoạn text lạ, có dấu hiệu băm nhỏ.
        *   **Hàm nguy hiểm:** `Shell`, `CreateObject`, `AutoOpen`, `Workbook_Open`.
        *   **Dấu vết mạng & Lệnh:** Địa chỉ URL, IP, câu lệnh thực thi PowerShell.
    3.  Hỗ trợ giải mã (deobfuscate) cấp độ cơ bản đối với chuỗi mã hóa Base64 hoặc chuỗi ghép nối.

#### Kết quả thực nghiệm với `olevba`:
<img width="940" height="601" alt="Kết quả phân tích mẫu từ MalwareBazaar bằng olevba" src="https://github.com/user-attachments/assets/98960512-f13e-47f2-aa9c-cf4192e826ce" />
*Figure 1: Trích xuất mã VBA của mẫu từ MalwareBazaar*

<img width="940" height="534" alt="Kết quả phân tích mẫu Excel GitHub bằng olevba - Ảnh 1" src="https://github.com/user-attachments/assets/bf894c38-25c5-4f41-90a8-c35824d4d4da" />
*Figure 2: Phân tích file Excel độc hại từ kho GitHub (Mẫu 1)*

<img width="940" height="549" alt="Kết quả phân tích mẫu Excel GitHub bằng olevba - Ảnh 2" src="https://github.com/user-attachments/assets/53974076-75a0-411c-badc-b594765656a4" />
*Figure 3: Phân tích file Excel độc hại từ kho GitHub (Mẫu 2)*

---

### 2.2. Công cụ mô phỏng động: `ViperMonkey`
*   **Nguồn:** [ViperMonkey GitHub](https://github.com/decalage2/ViperMonkey)
*   **Cách thức hoạt động:**
    *   Sử dụng cơ chế thông dịch (interpret) để parse và chạy giả lập toàn bộ luồng mã VBA mà không cần mở ứng dụng Microsoft Excel thật.
    *   **Theo dõi hành vi:** Giám sát liên tục sự thay đổi của biến, bám sát các logic rẽ nhánh (`If-Else`), vòng lặp (`For/While`).
    *   **Kết quả đầu ra:** Trả về các chuỗi ký tự sau khi giải mã, URL tải payload cuối cùng, và các câu lệnh PowerShell hoàn chỉnh được sinh ra trong bộ nhớ.

> **Ví dụ về khả năng deobfuscate:**
> Với đoạn macro làm rối: `Shell "powershell -enc aGVsbG8="`
> -> `ViperMonkey` sẽ phân tích luồng, giải mã đoạn Base64 để hiển thị câu lệnh tường minh mà không thực sự thực thi lệnh PowerShell đó ra hệ thống của bạn, đảm bảo an toàn tuyệt đối khi phân tích.

#### Kết quả thực nghiệm với `ViperMonkey`:
<img width="940" height="563" alt="Mô phỏng mẫu MalwareBazaar bằng ViperMonkey" src="https://github.com/user-attachments/assets/ff66fc53-56cf-420c-9acb-06ca45e46612" />
*Figure 4: Kết quả chạy giả lập mẫu VBA thu thập từ MalwareBazaar*

<img width="940" height="549" alt="Mô phỏng mẫu Excel GitHub bằng ViperMonkey - Ảnh 1" src="https://github.com/user-attachments/assets/cd65c412-c688-448d-b3f6-21596137d82f" />
*Figure 5: Luồng thực thi tệp Excel độc hại trên GitHub (Mẫu 1)*

<img width="940" height="617" alt="Mô phỏng mẫu Excel GitHub bằng ViperMonkey - Ảnh 2" src="https://github.com/user-attachments/assets/cef04f9c-51c4-4f65-aa1f-031117a3b457" />
*Figure 6: Luồng thực thi tệp Excel độc hại trên GitHub (Mẫu 2)*

---

## 3. Dữ liệu mẫu lành tính (Benign VBA)

Bộ dữ liệu đối chứng phục vụ cho việc huấn luyện/kiểm thử mô hình bao gồm **42 tệp tin dữ liệu/mã nguồn VBA sạch**.

**Chi tiết cấu trúc cây thư mục:**
```text
Thư mục gốc: Benign_VBA
├── 9 file đơn lẻ + 2 thư mục con
├── Thư mục: ScottSchaen-excel-vba-macros-master
│   └── 12 file mã nguồn mở rộng (.bas)
└── Thư mục: VBA-Code-Library-master
    └── 20 file mã nguồn thư viện tiện ích
```

## 4. Dữ liệu cập nhật từ MalwareBazaar Daily

Nguồn cấp: [MalwareBazaar Daily Datalake](https://datalake.abuse.ch/malware-bazaar/daily/)

Thời gian thu thập: Từ tháng 04/2025 đến tháng 05/2026

Tổng số lượng: 635 tệp tin định dạng nâng cao (.docm, .xlsm, .ppam) được xác nhận có chứa mã độc mã hóa/nhúng trong cấu trúc VBA.
