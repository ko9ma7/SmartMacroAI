# SmartMacroAI

<div align="center">

![Version](https://img.shields.io/badge/version-1.5.6-blue)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2B-lightgrey)
![.NET](https://img.shields.io/badge/.NET-8.0-purple)
![License](https://img.shields.io/badge/license-MIT-green)
![Language](https://img.shields.io/badge/language-VI%20%7C%20EN-orange)

**[English](#english) | [Tiếng Việt](#tiếng-việt)**

*Created by Phạm Duy – Giải pháp tự động hóa thông minh.*

</div>

---

## English

### What is SmartMacroAI?

SmartMacroAI is a powerful Windows macro automation tool with AI integration. Automate repetitive tasks, game farming, web form filling, and more — with support for anti-cheat games via kernel-level Interception driver.

### Key Features

- **Multi-mode Click** — Stealth (PostMessage), SendInput, Raw Input, Driver Level (Interception)
- **Key Press** — PostMessage, SendInput, Raw/ScanCode, Driver Level
- **Mouse Button Support** — Left, Right, Middle click across all modes
- **Anti-Cheat Support** — Kernel driver bypass (MapleStory, Cabal Online, etc.)
- **Image Recognition** — Find image on screen, auto click at found position
- **OCR Text Detection** — Read text from screen regions via Tesseract
- **CSV Auto Fill** — Loop CSV rows into web/desktop forms
- **AI Integration** — OpenAI & Gemini support for smart decision-making
- **Multi Dashboard** — Monitor and control multiple macros simultaneously
- **Bilingual UI** — English & Vietnamese with 700+ localization keys
- **Macro Lock** — Password-protect your macros
- **Scheduler** — Run macros at specific times or on startup
- **Visual Coordinate Picker** — Click to pick coordinates on screen
- **Run History** — Track macro execution logs and statistics
- **Telegram Notifications** — Get notified when macros complete
- **Script Sharing** — Share macros via encoded strings
- **6 Macro Templates** — Ready-to-use templates with full action chains
- **Anti-Detection** — Human-like mouse movement and random delays
- **Web Automation** — Playwright-based browser control

### System Requirements

| Component | Minimum |
|-----------|---------|
| OS | Windows 10 x64 or later |
| .NET | .NET 8.0 Runtime |
| RAM | 4 GB |
| Storage | 200 MB |

### Quick Start

1. Download the latest installer from [Releases](https://github.com/TroniePh/SmartMacroAI/releases)
2. Run `SmartMacroAI_Setup.exe` as Administrator
3. Launch SmartMacroAI
4. Click **+ New Macro** — choose a template or start from scratch
5. Set your target window — add actions — press **Run**

### Driver Level Mode (Anti-Cheat Games)

> Allows macros to work inside games protected by anti-cheat systems.

1. Go to **Settings → Driver Level**
2. Click **Install Now** — approve the UAC prompt
3. Restart your PC
4. Re-open SmartMacroAI — select **Driver Level** mode in any click/key action
5. Click **Test Driver** to verify the installation

**Supported games:** MapleStory, Cabal Online, and most DirectX games with anti-cheat.

### Template Gallery

| Template | Actions | Category |
|----------|---------|----------|
| Auto Login Website | 7 actions | Web |
| Auto Fill CSV Data | 9 actions (loop) | Web |
| Auto Repeat Task | 3 actions (loop) | Desktop |
| Image Detection Loop | 4 actions (loop) | Desktop |
| Hotkey Automation | 3 actions (loop) | General |
| Game Farming Loop | 6 actions (loop) | Desktop |

### Build from Source

```bash
# Requirements: .NET 8 SDK, Visual Studio 2022 or VS Code
git clone https://github.com/TroniePh/SmartMacroAI.git
cd SmartMacroAI
dotnet restore
dotnet build
dotnet run
```

### Build Installer

```bash
dotnet publish -c Release -r win-x64 --self-contained false -o publish
# Requires Inno Setup 6
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\SmartMacroAI_Setup.iss
```

### Changelog

See [CHANGELOG.md](CHANGELOG.md) for full version history.

### Contributing

Pull requests are welcome! Please open an issue first to discuss proposed changes.

### Donate / Support the Developer

If you find SmartMacroAI useful, consider supporting the developer to maintain and grow the project:

**PayPal:**
[![Donate via PayPal](https://img.shields.io/badge/Donate-PayPal-00457C?logo=paypal)](https://www.paypal.com/paypalme/nhocbobi22)

- **PayPal Email:** nhocbobi22@gmail.com

### Support & Contact

- **GitHub Issues:** [Report Bugs](https://github.com/TroniePh/SmartMacroAI/issues)
- **Website:** [SmartMacroAI Landing Page](https://tronieph.github.io/SmartMacroAI-Website/)

---

## Tiếng Việt

### SmartMacroAI là gì?

SmartMacroAI là phần mềm macro tự động hóa mạnh mẽ cho Windows, tích hợp AI. Tự động hóa các tác vụ lặp lại, farm game, điền form web, v.v. — với hỗ trợ game anti-cheat qua driver Interception cấp nhân.

### Tính năng nổi bật

- **Click đa chế độ** — Stealth (PostMessage), SendInput, Raw Input, Driver Level (Interception)
- **Nhấn phím** — PostMessage, SendInput, Raw/ScanCode, Driver Level
- **Hỗ trợ nút chuột** — Trái, Phải, Giữa trên tất cả chế độ
- **Hỗ trợ Anti-Cheat** — Bypass driver nhân (MapleStory, Cabal Online...)
- **Nhận diện hình ảnh** — Tìm ảnh trên màn hình, tự click vào vị trí tìm thấy
- **OCR nhận dạng chữ** — Đọc text từ vùng màn hình qua Tesseract
- **Tự động điền CSV** — Lặp dữ liệu CSV vào form web/desktop
- **Tích hợp AI** — Hỗ trợ OpenAI & Gemini cho quyết định thông minh
- **Multi Dashboard** — Theo dõi và điều khiển nhiều macro cùng lúc
- **Giao diện song ngữ** — Tiếng Anh & Tiếng Việt với hơn 700 khóa nội dung
- **Khóa Macro** — Bảo vệ macro bằng mật khẩu
- **Lịch chạy** — Đặt lịch chạy macro theo giờ hoặc khi khởi động
- **Chọn tọa độ trực quan** — Click để lấy tọa độ trên màn hình
- **Lịch sử chạy** — Theo dõi log và thống kê thực thi macro
- **Thông báo Telegram** — Nhận thông báo khi macro hoàn thành
- **Chia sẻ Script** — Chia sẻ macro qua chuỗi mã hóa
- **6 Mẫu Macro** — Sẵn sàng sử dụng với chuỗi action đầy đủ
- **Chống phát hiện** — Di chuyển chuột giống người và delay ngẫu nhiên
- **Tự động hóa Web** — Điều khiển trình duyệt qua Playwright

### Yêu cầu hệ thống

| Thành phần | Tối thiểu |
|-----------|---------|
| Hệ điều hành | Windows 10 x64 trở lên |
| .NET | .NET 8.0 Runtime |
| RAM | 4 GB |
| Lưu trữ | 200 MB |

### Bắt đầu nhanh

1. Tải installer mới nhất tại [Releases](https://github.com/TroniePh/SmartMacroAI/releases)
2. Chạy `SmartMacroAI_Setup.exe` với quyền Administrator
3. Mở SmartMacroAI
4. Bấm **+ Macro mới** — chọn mẫu hoặc tạo từ đầu
5. Chọn cửa sổ mục tiêu — thêm action — bấm **Chạy**

### Driver Level Mode (Game Anti-Cheat)

> Cho phép macro hoạt động trong game có hệ thống chống gian lận.

1. Vào **Cài đặt → Driver Level**
2. Bấm **Cài đặt ngay** — xác nhận UAC
3. Khởi động lại máy tính
4. Mở lại SmartMacroAI — chọn chế độ **Driver Level** trong action click/phím
5. Bấm **Test Driver** để kiểm tra cài đặt

**Game hỗ trợ:** MapleStory, Cabal Online và hầu hết game DirectX có anti-cheat.

### Ủng hộ / Donate

Nếu bạn thấy SmartMacroAI hữu ích, hãy ủng hộ tác giả để duy trì và phát triển dự án:

**PayPal:**
[![Donate via PayPal](https://img.shields.io/badge/Donate-PayPal-00457C?logo=paypal)](https://www.paypal.com/paypalme/nhocbobi22)

- **PayPal Email:** nhocbobi22@gmail.com

**QR Bank (Vietnam):**
<!-- Thêm hình QR bank vào đây nếu có -->
Xem mã QR chuyển khoản trong phần **About** của ứng dụng.

### Báo lỗi & Hỗ trợ

- **GitHub Issues:** [Báo lỗi tại đây](https://github.com/TroniePh/SmartMacroAI/issues)
- **Website:** [Trang giới thiệu SmartMacroAI](https://tronieph.github.io/SmartMacroAI-Website/)

---

<div align="center">

Made with ❤️ in Vietnam

*Created by Phạm Duy – Giải pháp tự động hóa thông minh.*

</div>
