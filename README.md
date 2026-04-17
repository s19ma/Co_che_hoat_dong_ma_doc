https://github.com/s19ma/Co_che_hoat_dong_ma_doc/tree/main/All_Excel_Doc_Macro

**1. Số lượng và Phân loại**
Tổng cộng có 17 mẫu thử nghiệm mã độc và 2 mẫu sạch được liệt kê

- Nhóm 1 (4 mẫu): Sử dụng hàm EXEC để thực thi các lệnh shell (như khởi chạy powershell.exe với lệnh đã mã hóa)

- Nhóm 2 (4 mẫu): Sử dụng hàm REGISTER để gọi các hàm từ thư viện Kernel32.dll (như VirtualAlloc, WriteProcessMemory, CreateThread) nhằm ghi mã độc trực tiếp vào bộ nhớ và thực thi, thường làm Excel bị treo sau khi chạy

- Nhóm 3 (5 mẫu): Sử dụng hàm CALL để truy cập urlmon.dll tải tệp từ bên ngoài (như từ GitHub) và thực thi chúng qua shell32.dll

- Nhóm 4 (4 mẫu): Các mẫu nâng cao kết hợp kỹ thuật từ các nhóm trên nhưng loại bỏ cơ chế tự động chạy (Auto_Open) và thêm các kỹ thuật phát hiện môi trường ảo (sandboxing) để né tránh phân tích


**2. Các kỹ thuật lẩn trốn (Evasion Techniques)**
- Làm rối mã (Obfuscation):
  - Sử dụng các hàm Excel như CHAR, MID, CODE, HEX2DEC để che giấu chuỗi lệnh độc hại
  - Dùng hàm FORMULA để xây dựng nội dung mã độc một cách linh hoạt trong thời gian thực khi tệp được mở
  - Sử dụng vòng lặp while để tạo động các hàm thực thi

- Ẩn giấu thành phần (Hidden Sheets):
  - Sử dụng các bảng tính ở chế độ "Hidden" (Ẩn) hoặc "VeryHidden" (Rất ẩn) để người dùng thông thường không thể nhìn thấy mã macro.
  - Tài liệu lưu ý rằng việc sử dụng "VeryHidden" làm tăng đáng kể khả năng bị các phần mềm diệt virus phát hiện

- Phát hiện môi trường ảo/Sandbox (Sandboxing Detection):
 - Sử dụng các hàm GET.WORKSPACE, GET.DOCUMENT, GET.WINDOW để kiểm tra thông tin hệ thống
 - Kiểm tra điều kiện môi trường: Mã độc chỉ thực thi nếu phát hiện có chuột, cửa sổ Excel được phóng to tối đa, máy tính có khả năng ghi âm/phát thanh, hoặc các đường dẫn t


<img width="1536" height="2752" alt="image" src="https://github.com/user-attachments/assets/b75e7cff-4593-4270-97ab-b02274941f7e" />

<img width="1553" height="645" alt="image" src="https://github.com/user-attachments/assets/ca86ebb0-3de9-4489-9980-c7500af580e9" />
