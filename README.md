## Table of Contents
[Tổng quan](#tổng-quan-dataset)

[Tool](#tool-extract-vba)

## Tổng quan dataset

https://github.com/s19ma/Co_che_hoat_dong_ma_doc/tree/main/All_Excel_Doc_Macro

**1. Số lượng và Phân loại**
Tổng cộng có 17 mẫu thử nghiệm mã độc và 2 mẫu sạch được liệt kê

- Nhóm 1 (4 mẫu): Sử dụng hàm EXEC để thực thi các lệnh shell (như khởi chạy powershell.exe với lệnh đã mã hóa)

- Nhóm 2 (4 mẫu): Sử dụng hàm REGISTER để gọi các hàm từ thư viện Kernel32.dll (như VirtualAlloc, WriteProcessMemory, CreateThread) nhằm ghi mã độc trực tiếp vào bộ nhớ và thực thi, thường làm Excel bị treo sau khi chạy

- Nhóm 3 (5 mẫu): Sử dụng hàm CALL để truy cập urlmon.dll tải tệp từ bên ngoài (như từ GitHub) và thực thi chúng qua shell32.dll

- Nhóm 4 (4 mẫu): Các mẫu nâng cao kết hợp kỹ thuật từ các nhóm trên nhưng loại bỏ cơ chế tự động chạy (Auto_Open) và thêm các kỹ thuật phát hiện môi trường ảo (sandboxing) để né tránh phân tích

- Nhóm 5 (16 mẫu) – Chưa phân loại: Các mẫu VBA phức tạp hoặc bị obfuscate mạnh, chưa xác định rõ hành vi chính.

- Nhóm 7 (78 mẫu): Các mẫu VBA từ [MalwareBazaar](https://bazaar.abuse.ch/browse/tag/vba/)

  
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
 - Kiểm tra điều kiện môi trường: Mã độc chỉ thực thi nếu phát hiện có chuột, cửa sổ Excel được phóng to tối đa, máy tính có khả năng ghi âm/phát thanh.


<img width="1536" height="2752" alt="image" src="https://github.com/user-attachments/assets/b75e7cff-4593-4270-97ab-b02274941f7e" />

<img width="1553" height="645" alt="image" src="https://github.com/user-attachments/assets/ca86ebb0-3de9-4489-9980-c7500af580e9" />

## Tool-extract-vba 
Tool [olevba](https://github.com/decalage2/oletools/wiki/olevba) Phân tích tĩnh

 <img width="940" height="601" alt="image" src="https://github.com/user-attachments/assets/98960512-f13e-47f2-aa9c-cf4192e826ce" />
 
Figure 1 VBA từ MalwareBazaar


 <img width="940" height="534" alt="image" src="https://github.com/user-attachments/assets/bf894c38-25c5-4f41-90a8-c35824d4d4da" />
 
Figure 2 File Excel từ github


 <img width="940" height="549" alt="image" src="https://github.com/user-attachments/assets/53974076-75a0-411c-badc-b594765656a4" />
 
Figure 3 File Excel từ github


[ViperMonkey](https://github.com/decalage2/ViperMonkey) Phân tích động dạng giả lập (emulation)


<img width="940" height="563" alt="image" src="https://github.com/user-attachments/assets/ff66fc53-56cf-420c-9acb-06ca45e46612" />

Figure 4 VBA từ MalwareBazaar

<img width="940" height="549" alt="image" src="https://github.com/user-attachments/assets/cd65c412-c688-448d-b3f6-21596137d82f" />

Figure 5 File Excel từ github

<img width="940" height="617" alt="image" src="https://github.com/user-attachments/assets/cef04f9c-51c4-4f65-aa1f-031117a3b457" />

Figure 6 File Excel từ github
