# LAN Share — Tool_Share_LAN

Bộ công cụ triển khai **tập trung thư mục / ổ đĩa dùng chung và máy in** cho nhiều máy tính trong cùng mạng LAN (workgroup), không cần Active Directory / domain. Gồm một **máy chủ giao diện Windows** (PowerShell + WinForms) và các **gói đăng ký máy con** cho Windows và macOS.

> Máy chủ điều khiển: `bin\May-chu.ps1` — khởi chạy qua `Mo-may-chu.cmd` hoặc `Mo-may-chu.vbs` (tự yêu cầu quyền Administrator, chạy ẩn).

---

## 1. Tính năng chính

| Nhóm | Chức năng |
|---|---|
| Khởi tạo máy chủ | Chuyển mạng Public → Private, tắt Windows Firewall (3 profile), tạo/cập nhật tài khoản share, khởi động `LanmanServer`, mở bộ nhận đăng ký trên cổng **41866**. Lưu bản sao cấu hình mạng để có thể khôi phục. |
| Tài nguyên | Thêm thư mục / gốc ổ đĩa hoặc chọn máy in (giao diện có hình); nạp các share có sẵn trên Windows. Quyền: **Keep** (giữ nguyên), **Read / Change / Full** (cấp lại cả Share + NTFS). Chọn ổ máy con **Auto** (Z→D) hoặc cố định **D–Z**. |
| Triển khai máy con | Tick máy → xếp lệnh → máy con tự map ổ đĩa (`WNetAddConnection2`) hoặc thêm máy in. Thử lần lượt Administrator → Guest → tài khoản share, tự lưu credential. Có hàng đợi lệnh, tự sửa kết nối thiếu (`Desired` / repair). |
| Máy in | Bật share máy in, cấp quyền in; đóng gói driver INF đã ký số từ máy chủ, cài sẵn bằng `pnputil` + `InstallPrinterDriverFromPackageW`; bộ hỗ trợ chạy dưới quyền SYSTEM (Scheduled Task) để chuẩn bị driver; fallback mở hộp thoại `printui.dll` trong phiên người dùng. |
| Phân quyền bộ phận | Mỗi bộ phận có tài khoản local riêng trên máy chủ. Gán máy vào bộ phận; mỗi tài nguyên đặt quyền theo từng bộ phận (Read/Change/Full hoặc In/Không). Áp dụng khi triển khai, có sao lưu ACL/share trước khi đổi. |
| Máy mới | Khi có máy đăng ký lần đầu, hỏi dùng lại "rule" từ máy đã triển khai thành công, hoặc cấu hình mới. Có chế độ "máy này thay máy cũ cài lại Windows". |
| Tiến trình & báo cáo | Xem tiến trình theo từng máy (host) và cửa sổ tiến trình trên máy con; xuất báo cáo **HTML** (đọc/in) hoặc **JSON** (đối chiếu) gồm trạng thái Windows thực tế, cấu hình, bộ phận, máy con, lịch sử triển khai (đã ẩn mật khẩu). |
| Sao lưu / phục hồi | File `.lanbackup` mã hóa AES bằng mật khẩu (≥10 ký tự, PBKDF2 150.000 vòng).Khôi phục trước khi khởi tạo lại; tự hủy các lệnh Pending/Running cũ. Cập nhật IP / tạo lại gói đăng ký khi đổi mạng. |
| Bảo trì | Khôi phục mạng/firewall, hoàn tác bản sửa SID tạm, khởi động lại máy chủ. |
| Cập nhật | Máy chủ tự kiểm tra GitHub khi mở tool. Chỉ tải và cài khi người dùng xác nhận; giữ dữ liệu máy chủ. |
| Tương thích | Host: Windows 10/11 (PowerShell 5.1). Client Windows 10/11 đầy đủ tính năng; **Windows 7 SP1 + WMF 5.1 + .NET 4.6.2+** dùng bộ adapter legacy. Client **macOS** (Python 3.10+): nhận thư mục SMB trong Finder, chưa hỗ trợ máy in Windows. |
| Xử lý SID trùng | Tự đối chiếu machine SID và sự kiện `LsaSrv 6167`, áp bản sửa tạm (Registry override `1517186191`) trên Windows 11 build 26100/26200, kèm nút hoàn tác. |

---

## 2. Cấu trúc thư mục

```text
Tool_Share_LAN/
├── Mo-may-chu.cmd            # Launcher CMD → chạy May-chu.ps1 với UAC (RunAs)
├── Mo-may-chu.vbs            # Launcher VBS tương đương (ShellExecute "runas")
├── readme.md                 # Tài liệu này
├── huongdansudung.html       # Hướng dẫn sử dụng chi tiết (HTML)
└── bin/
    ├── May-chu.ps1           # GUI máy chủ + điều phối toàn bộ luồng (file chính, ~1000 dòng)
    ├── Common.ps1            # Tiện ích chung: JSON, crypto AES/HMAC/RSA, SID, mạng, P/Invoke Win32
    ├── Host-Core.ps1         # Trạng thái/cấu hình máy chủ, giải mã yêu cầu, share/print, đóng gói client .cmd
    ├── Client.ps1            # Agent máy con: poll lệnh, map ổ / thêm máy in, fallback credential, SID
    ├── Install-Client.ps1    # Cài đặt & đăng ký máy con Windows; khởi động agent nền
    ├── Network-Setup.ps1     # Apply/Restore cấu hình mạng + cài printer helper + stage driver (UAC)
    ├── Printer-Helper.ps1    # Worker SYSTEM chuẩn bị driver máy in (Scheduled Task)
    ├── Driver-Packages.ps1   # Xuất/gói/ký driver INF; cài driver bằng package API
    ├── Print-Preferences.ps1 # Chuẩn bị các tùy chọn in trên máy con Windows 11
    ├── Update.ps1            # Kiểm tra, xác minh và cài bản phát hành GitHub
    ├── Departments.ps1       # Logic bộ phận & phân quyền share/NTFS/printer
    ├── Departments-UI.ps1    # 3 dialog: quản lý bộ phận, quyền bộ phận, gán bộ phận
    ├── New-Clients.ps1       # Dùng lại rule từ máy đã triển khai thành công
    ├── Printer-Picker.ps1    # Dialog chọn máy in dạng ô có icon
    ├── Reports.ps1           # Tạo/ xuất báo cáo HTML, JSON
    ├── Recovery.ps1          # Backup/restore .lanbackup có mật khẩu
    ├── Client-Progress.ps1   # Cửa sổ tiến trình trên máy con + gửi snapshot về host
    ├── Host-Progress.ps1     # Cửa sổ tiến trình máy con trên host
    ├── Display.ps1           # Giới hạn cửa sổ theo màn hình, viewport cuộn, dialog DPI
    ├── Layout.ps1            # Bố cục 3 nhóm + thanh công cụ cho GUI máy chủ
    ├── Theme.ps1             # Màu sắc / style nút, bảng, log
    ├── Legacy-Client.ps1     # Adapter COM cho Windows 7 (mạng, firewall, ổ, máy in)
    └── mac/
        ├── Dang-ky-macOS.command # Script đăng ký Mac (venv + cryptography)
        ├── Go-dang-ky.command    # Script gỡ tự khởi động Mac
        └── mac_client.py         # Agent Python: đăng ký, poll, mount SMB qua Finder/osascript
```

---

## 3. Kiến trúc & luồng hoạt động

```text
 [Máy chủ - May-chu.ps1]                         [Máy con Windows]              [Mac]
  Khởi tạo (mạng/account/firewall)                  Chạy Dang-ky-may-con.cmd     Chạy .command
  HTTP listener  http://IP:41866/api/  <── AES/HMAC ── agent nền (poll 5s) ──>  agent Python (poll 5s)
  Chọn tài nguyên + quyền + bộ phận                 map ổ D-Z / Add-Printer     mount SMB trong Finder
  "Triển khai" → ghi jobs.json                      báo kết quả/heartbeat        báo kết quả
  Nhận Progress/Result → cập nhật bảng              lưu Ket-qua-trien-khai.txt  lưu Ket-qua-*.txt
```

- **Giao tiếp**: HTTP POST tới `/api/` trên cổng 41866. Mỗi bản tin là envelope JSON gồm `Body` (JSON nén → AES-256-CBC base64) + `Signature` (HMAC-SHA256). Khi đăng ký, khóa phiên 32 byte của máy con được bọc bằng RSA-2048/OAEP của máy chủ; máy chủ xác minh `RegistrationProof`. Mọi yêu cầu có `Time` (cho phép lệch ≤ 2 phút) và `Nonce` (chống phát lại, cache 5 phút).
- **Danh tính máy con**: `ClientId` (GUID 32 hex) + `IdentityKey` + `DeviceIdentity` (hash của MachineGuid/UUID máy). Gói đăng ký cũ tự giữ lại `ClientId` khi chạy lại.
- **Dữ liệu máy chủ** tại `%ProgramData%\ToolShareLAN\Host\`: `host.json`, `clients.json`, `client-keys.json`, `jobs.json`, `resources.json`, `departments.json`, `network-before.json`, `sid-before.json`, `Backups\`. Thư mục bị khóa ACL chỉ cho SYSTEM + Administrators.
- **Dữ liệu máy con Windows** tại `%LOCALAPPDATA%\LanShare\<SiteId>\`: `client.json`, `client.log`, `results.json`, `deployment-progress.json`, `Dang-ky-report.txt`, `Ket-qua-trien-khai.txt`, `Chan-doan-driver.txt`, `Doi-chieu-driver.txt`, marker `stop`. Startup: `LanShare-<SiteId>.vbs`. Bộ hỗ trợ in: `%ProgramData%\LanShare\<SiteId>\PrinterHelper\<ClientId>\`.
- **Dữ liệu Mac**: `~/Library/Application Support/LanShare/<SiteId>/` (client.json, results.json, mounts.json, agent.log, mac_client.py) + `~/Library/LaunchAgents/local.lanshare.<SiteId>.plist`.

---

## 4. Bắt đầu nhanh

1. Giải nén giữ nguyên thư mục `Tool_Share_LAN`. Trên máy chủ, nhấp đúp `Mo-may-chu.cmd` (hoặc `.vbs`) → chấp nhận UAC.
2. Chọn **IP máy chủ**, giữ `User share = print` / mật khẩu (đổi nếu cần), bấm **Khởi tạo máy chủ**.
3. Ở bảng tài nguyên: bấm **Thêm thư mục / ổ đĩa** hoặc **Chọn máy in** (hoặc **Nạp share có sẵn**), tick vào cột **Chọn**, đặt quyền và ổ máy con.
4. Bấm **Tạo file đăng ký** → chọn **Windows (.cmd)** hoặc **macOS (.zip)**; copy file cho các máy con và chạy (Windows: nhấp đúp, chấp nhận UAC; Mac: giải nén, chạy `Dang-ky-macOS.command`, cần Python 3.10+ và Internet lần đầu).
5. Khi máy xuất hiện trong bảng "Máy con", tick chọn (có thể **Gán bộ phận**), bấm **Triển khai đã chọn**. Nhấp đúp một dòng để xem tiến trình.

📖 **Hướng dẫn chi tiết từng nút, phân quyền, sao lưu, báo cáo và khắc phục sự cố:** mở [`huongdansudung.html`](huongdansudung.html).

---

## 5. Cấu hình mặc định

| Tham số | Giá trị |
|---|---|
| Cổng bộ nhận máy con | `41866` (HTTP) |
| Tài khoản share mặc định | `print` / `12345678` |
| Mật khẩu gợi ý cho bộ phận | `12345678` |
| Quyền mặc định khi thêm thư mục | `Change`, ổ `Auto` |
| Quyền mặc định khi thêm máy in | `Read` (quyền in) |
| Ký tự ổ tự động | từ **Z** xuống **D** |
| Thời gian coi máy "Offline" | 35 giây không có phản hồi |
| Tần suất poll của máy con | 5 giây |
| Gói driver tối đa | 256 MB giải nén / 48 MB nén (mỗi gói) |
| Yêu cầu mật khẩu sao lưu | tối thiểu 10 ký tự |

---

## 6. Lưu ý bảo mật & giới hạn

- ⚠️ **Khởi tạo máy chủ sẽ tắt Windows Firewall ở cả 3 profile và chuyển mạng Public sang Private.** Chỉ dùng trong LAN nội bộ tin cậy; luôn bấm **Khôi phục mạng** (trong "Bảo trì / Khôi phục mạng") khi không còn sử dụng. Tool chỉ hỗ trợ máy **workgroup**, từ chối chạy nếu máy thuộc domain.
- Mật khẩu share/admin và khóa được lưu trong cấu hình/gói đăng ký để máy con tự kết nối; hãy bảo vệ file `.cmd`/`.zip` và file `.lanbackup` bằng mật khẩu mạnh. Báo cáo đã tự động thay mật khẩu bằng `[ẩn]`.
- Máy con Windows 10/11 mới cài được driver tự động; Windows 7 chỉ nhận thư mục/máy in ở mức cơ bản. **macOS chỉ nhận thư mục SMB**, không map ký tự ổ, không tự cài máy in Windows, không đổi firewall, không xử lý SID.
- **Gói Windows có tick máy in:** trên Windows 11, bước đăng ký chuẩn bị dùng driver hãng bằng cách tắt Windows protected print mode (nút **Set up** khi không bị chính sách khóa), **Default install printers using Windows Ready Print** nếu bản Windows hỗ trợ, **Let Windows manage my default printer** cho tài khoản đăng ký, và **Download drivers and device software over metered connections**. Việc tắt Protected Print làm giảm một lớp bảo vệ; không dùng gói này trên máy cần giữ chế độ đó. Công cụ không đổi chính sách đang bắt buộc bật chế độ in, không bỏ kiểm tra chữ ký driver, và không tự bật lại các tùy chọn khi gỡ đăng ký.
- Phân quyền bộ phận đặt lại ACL của thư mục được chọn (chỉ giữ Administrators, SYSTEM và các bộ phận được chọn); không áp được cho gốc ổ đĩa và không tự sửa ACL của thư mục con đã có quyền riêng. Luôn có bản sao ACL trong `Backups\` trước khi đổi.
- Bản sửa SID tạm chỉ áp cho Windows 11 build 26100/26200 khi có sự kiện `LsaSrv 6167`; cần khởi động lại và có nút **Hoàn tác SID tạm**.
- Đổi IP máy chủ: vào menu **Sao lưu / Phục hồi → Cập nhật IP máy chủ / tạo lại gói đăng ký**, chọn IP mới, khởi tạo lại và phát gói đăng ký mới cho máy con nếu DNS không phân giải được tên máy chủ.

---

## 7. Khắc phục sự cố nhanh

| Triệu chứng | Kiểm tra |
|---|---|
| Máy con không xuất hiện | Máy chủ đã "Khởi tạo" chưa? Đúng IP/cổng 41866? Hai máy ping được nhau? Xem `Dang-ky-report.txt` và `client.log` trên máy con. |
| Báo "lệch đồng hồ quá 2 phút" | Chỉnh lại giờ hệ thống hai máy (chênh ≤ 120 giây). |
| Map ổ báo "ổ đã được sử dụng" | Chọn `Auto` hoặc ngắt ổ cũ; máy con tự chọn ký tự trống D–Z. |
| Máy in báo lỗi driver | Mở `Chan-doan-driver.txt` / `Doi-chieu-driver.txt`; tạo lại gói đăng ký có tick máy in để kèm driver đúng kiến trúc, chấp nhận UAC khi đăng ký lại. |
| Gói không có driver cho máy in hoặc Windows chưa xác nhận kết nối | Máy con mở trình cài Windows nếu chưa có driver cài sẵn thành công. Hoàn tất **Install driver** / UAC nếu được hỏi; mã thoát `0` của PrintUI không có nghĩa là đã cài xong. Nếu Windows vẫn không cài được, tick đúng máy in trên máy chủ, xuất gói Windows `.cmd` mới, đăng ký lại và xem mục `Driver` trong `Dang-ky-report.txt` trước khi triển khai lại. |
| Windows 11 vẫn hiện **Turn off** hoặc một tùy chọn in còn **On** | Mở lại **Settings → Bluetooth & devices → Printers & scanners** và xem từng mục trong `Dang-ky-report.txt`. Nếu báo do chính sách, liên hệ quản trị viên. Nếu báo chưa nạp lại Print Spooler, khởi động lại Windows rồi đăng ký lại. Gói chỉ có thư mục không đổi các tùy chọn in. |
| Kết nối báo 1219 (nhiều phiên SMB) | Đóng file đang mở trên máy con, đăng xuất Windows rồi triển khai lại. |
| Máy con chỉ truy cập được sau khi bỏ yêu cầu ký SMB | Ưu tiên cấu hình máy chủ để hỗ trợ ký SMB và dùng tài khoản chia sẻ. Nếu cần tương thích với máy chủ cũ trong LAN tin cậy, trên máy chủ chọn **Tạo file đăng ký → Xuất đăng ký Windows - tương thích SMB cũ...** rồi chạy gói mới trên máy con và xác nhận UAC. Có thể tương đương bằng CMD: `"C:\duong-dan\Dang-ky-may-con.cmd" /smb-compat`. Xem dòng **Tương thích SMB** trong `Dang-ky-report.txt`. Thay đổi áp dụng cho toàn máy con; `/remove` khôi phục yêu cầu ký nếu tool đã thay đổi. |
| Cần gỡ máy con | Tick máy → **Xóa máy đã chọn**. Để cài lại, chạy `Dang-ky-may-con.cmd /remove` để dỡ tự khởi động và khôi phục mạng. |
| Khôi phục cấu hình | Đóng tool, mở lại, vào **Sao lưu / Phục hồi → Khôi phục bản sao lưu** trước khi bấm Khởi tạo. |

---

## 8. Ghi chú kỹ thuật

- Mã nguồn thuần PowerShell 5.1 + .NET Framework (WinForms), không dependency ngoài trên máy chủ. Máy Mac cần `cryptography==50.0.2` trong venv riêng (tự cài lần đầu).
- `May-chu.ps1` có chế độ `-ValidateOnly [-PreviewPath]` để kiểm tra giao diện (dùng cho phát triển/kiểm thử).
- Các script builder nhúng mã nguồn vào gói đăng ký dưới dạng JSON → GZip → Base64, bootstrap nằm trong file `.cmd` với ký tự `:LANSHARE_PAYLOAD` (giới hạn 8191 ký tự cho dòng lệnh CMD).
- Client build hiện tại: `2026.11.26-smb-08`. Máy con đã đăng ký cần chạy gói Windows `.cmd` mới để cập nhật.

## 9. Cập nhật qua GitHub

- **Người dùng:** Mở LAN Share trên máy chủ; tool tự kiểm tra bản mới. Hoặc bấm **Sao lưu / Phục hồi → Kiểm tra cập nhật GitHub...**. Nếu có bản mới, xác nhận để cài rồi mở lại tool. Khởi tạo máy chủ nếu cần tiếp tục nhận kết nối. Không cài trong lúc đang triển khai máy con. Cấu hình và danh sách máy con trong `%ProgramData%\ToolShareLAN\Host` không bị thay đổi.
- Người đang dùng bản cũ chưa có chức năng cập nhật cần tải bản có chức năng này về **một lần**. Nếu bản phát hành có thay đổi cho chương trình máy con, xuất và chạy lại file đăng ký mới trên các máy cần cập nhật.
- **Người phát hành:** Tăng số trong `version.txt`, ví dụ `0.8.0`. Đóng thư mục tool thành `Tool_Share_LAN.zip` (có thể để một thư mục `Tool_Share_LAN` bao ngoài). Tạo [GitHub Release](https://github.com/tuantran19912512/Toolsharelan/releases) chính thức với tag `v0.8.0`, đính kèm ZIP và phát hành. Không dùng ZIP `Source code` tự sinh hoặc chỉ cập nhật nhánh `main`. Không đưa gói đăng ký đã xuất, dữ liệu nội bộ hay bản sao lưu vào ZIP. Updater xác minh SHA-256 từ GitHub và kiểm tra phiên bản trong ZIP khớp tag.

---

*Cập nhật 25/11/2026: sửa lỗi triển khai máy in và bổ sung tự cấu hình tùy chọn in Windows 11.*
