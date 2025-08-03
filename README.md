# PrinterOne 🖨️

[![GitHub stars](https://img.shields.io/github/stars/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/network)
[![GitHub issues](https://img.shields.io/github/issues/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/issues)
[![GitHub license](https://img.shields.io/github/license/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/blob/master/LICENSE)
[![GitHub release](https://img.shields.io/github/v/release/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/releases)

A comprehensive TCP network print server with integrated GUI management and system tray support for Windows.

**Copyright (c) 2025 xtieume@gmail.com**  
**GitHub: https://github.com/xtieume/PrinterOne**

![PrinterOne Screenshot](assets/screenshot.png)

## ✨ Features

- 🖨️ **TCP Network Print Server** - Receives raw print data over TCP and sends to local printer
- 🖥️ **Modern GUI Interface** - Easy-to-use graphical interface with tabbed layout
- 🔧 **System Tray Integration** - Minimize to tray and control from system tray menu
- 🚀 **Windows Auto-Start** - Add to Windows startup for automatic server launch
- 🧪 **Built-in Test Client** - Test connections and send sample print jobs
- 📊 **Real-time Logging** - Comprehensive logging with emoji indicators for easy reading
- ⚙️ **Configuration Management** - Persistent settings stored in JSON format
- 🔥 **Port Management** - Automatic port conflict resolution
- 🌐 **Network Discovery** - Shows local IP addresses for easy client configuration

## 🎯 Use Cases

- **Network Printing**: Share local printers over TCP network
- **Raw Data Printing**: Print raw data from applications, scripts, or devices
- **Legacy System Integration**: Bridge old systems to modern printers
- **Print Server**: Centralized printing solution for small networks

## 📋 Requirements

- Windows 10/11
- Python 3.9+ (for development)
- Local printer installed

## 🚀 Quick Start

### Option 1: Download Pre-built Executable (Recommended)
1. Download the latest `PrinterOne.exe` from [Releases](https://github.com/xtieume/PrinterOne/releases)
2. Run `PrinterOne.exe gui`
3. Select your printer and configure settings
4. Click "Start Server"

### Option 2: Build from Source
```bash
# Clone the repository
git clone https://github.com/xtieume/PrinterOne.git
cd PrinterOne

# Install dependencies
pip install -r requirements.txt

# Build executable
python build.py

# Run the executable
dist/PrinterOne.exe gui
```

## 🖥️ Usage Modes

### GUI Mode (Recommended)
```bash
PrinterOne.exe gui
```
- Full graphical interface
- System tray integration
- Real-time logging
- Built-in test client

### Console Mode
```bash
PrinterOne.exe
```
- Command-line interface
- Suitable for servers or automated deployments

### Test Client Mode
```bash
PrinterOne.exe test
```
- Interactive connection testing
- Send test print jobs

## ⚙️ Configuration

The application uses `config.json` for persistent settings:

```json
{
    "printer_name": "Canon MF220 Series",
    "port": 9100,
    "auto_start": false,
    "service_name": "PrinterOne",
    "service_description": "PrinterOne - Network print server for raw print data",
    "manual": true,
    "minimize_to_tray": true
}
```

### Configuration Options

| Option | Type | Description |
|--------|------|-------------|
| `printer_name` | string | Target printer name (from Windows printer list) |
| `port` | integer | TCP port to listen on (default: 9100) |
| `auto_start` | boolean | Start server automatically |
| `minimize_to_tray` | boolean | Minimize to system tray when closing |

## 🌐 Network Usage

Once the server is running, other machines can send print jobs to:
```
YOUR_IP_ADDRESS:9100
```

Example using netcat:
```bash
echo "Hello World" | nc YOUR_IP_ADDRESS 9100
```

Example using Python:
```python
import socket

def send_to_printer(data, host='192.168.1.100', port=9100):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect((host, port))
        s.send(data.encode())

send_to_printer("Test print job")
```

## 🔧 Advanced Features

### System Tray Integration
- Right-click tray icon for quick actions
- Start/stop server from tray menu
- Show/hide main window
- Minimize to tray on close

### Auto-Start Configuration
- Add to Windows startup registry
- Auto-start with configured printer
- Background operation support

### Logging System
- Real-time logging with emoji indicators
- Automatic log rotation (30-day retention)
- Separate test client logging
- File-based logging for debugging

## 🛠️ Development

### Project Structure
```
PrinterOne/
├── server.py          # Main application
├── build.py           # Build script
├── requirements.txt   # Python dependencies
├── config.json        # Configuration file
├── PrinterOne.spec    # PyInstaller spec
├── assets/           # Assets folder
│   └── screenshot.png # Application screenshot
├── dist/             # Built executables
└── logs/             # Application logs
```

### Building
```bash
# Install development dependencies
pip install -r requirements.txt

# Build executable
python build.py

# Clean build
python build.py --clean
```

### Dependencies
- `pywin32` - Windows API integration
- `pystray` - System tray support
- `Pillow` - Image processing
- `psutil` - Process management
- `pyinstaller` - Executable packaging

## 🐛 Troubleshooting

### Common Issues

**Q: Tray icon not visible**  
A: Check Windows notification area settings to show PrinterOne icon.

**Q: Port already in use**  
A: The server automatically kills conflicting processes. Check firewall settings.

**Q: Printer not found**  
A: Ensure the printer is installed and accessible from Windows. Use exact printer name from Windows printer list.

### Debug Mode
Enable verbose logging by running:
```bash
PrinterOne.exe gui --debug
```

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the project
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📞 Support

- 🐛 **Bug Reports**: [GitHub Issues](https://github.com/xtieume/PrinterOne/issues)
- 💡 **Feature Requests**: [GitHub Issues](https://github.com/xtieume/PrinterOne/issues)
- 📧 **Email**: xtieume@gmail.com

## ⭐ Show Your Support

If this project helped you, please consider giving it a ⭐ on GitHub!

---

<div align="center">
Made with ❤️ by <a href="mailto:xtieume@gmail.com">xtieume</a>
</div>