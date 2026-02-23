# WindowsSmartTaskbar 🚀

A modern, sleek, and high-performance Windows utility designed to streamline your workflow by organizing your favorite applications into custom categories. Access everything instantly via an interactive system tray menu.

---

## ✨ Key Features

- **📂 Smart Categorization**: Organize your tools, games, and apps into logical groups.
- **⚡ Quick-Launch Menu**: Left-click the tray icon to access a categorized mini-launcher instantly.
- **🎯 Full Management**: Double-click to open the main dashboard for adding, editing, or removing programs.
- **🌍 Multi-Language Support**: Fully localized in **English (Default)**, **Swedish**, and **Turkish**.
- **🛠️ Zero-Config Persistence**: Your categories and program lists are automatically saved and restored.
- **🔄 Windows Autostart**: Option to launch automatically when you start your PC.
- **🎨 Modern Aesthetics**: A clean, professional UI with custom-drawn iconography and a sleek blue design.

---

## 🚀 Getting Started

### Prerequisites

- **Windows 10/11**
- **.NET 8.0 Runtime** (or later)

### Installation

1. **Download** the latest release or clone the repository.
2. **Build** the project using Visual Studio or the terminal:
   ```powershell
   dotnet build
   ```
3. **Launch** the executable found in `bin/Debug/net8.0-windows/WindowsSmartTaskbar.exe`.

---

## 📖 How to Use

### 1. The System Tray (Your Hub)
- **Left-Click**: Opens the **Quick-Launch Menu**. Hover over a category and click a program to start it.
- **Double-Click**: Transitions to the **Main Dashboard** for full management.
- **Right-Click**: Opens **Settings**, where you can change language, toggle autostart, or reset data.

### 2. Managing Programs
- Use the **Add Program** button to select an `.exe` or `.lnk` (shortcut).
- **Edit Name**: Keep your list clean by renaming shortcuts.
- **Categories**: Create custom categories and move programs between them to stay organized.

---

## 🛠️ Technical Overview

- **Framework**: .NET 8.0 (Windows Forms)
- **Language**: C# 12
- **Data Engine**: JSON-based storage (fast & lightweight)
- **Iconography**: Custom programmatic GDI+ rendering
- **API Integration**: Win32 API for advanced focus and window management

---

## 📜 License

Distributed under the **MIT License**. See `LICENSE` for more information.

---

## 👨‍💻 Author

**Created 2026 by © nRn World**

📧 bynrnworld@gmail.com

## 🙏 Support

If you like this project, consider to:

⭐ **Star the project** on GitHub  
☕ **Buy me a coffee** [buymeacoffee.com/nrnworld](https://buymeacoffee.com/nrnworld)  
📢 **Share with your friends**
