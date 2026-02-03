# NearLock

**Automatic PC lock when your Bluetooth device disconnects**

NearLock monitors a paired Bluetooth device (like your phone) and automatically locks your Windows PC when you walk away. When you return and your device reconnects, you simply unlock your PC as usual.

![Windows](https://img.shields.io/badge/platform-Windows-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **Bluetooth Classic & BLE detection** - Works with both classic Bluetooth and Bluetooth Low Energy devices
- **System tray application** - Runs quietly in the background with status indicator
- **First-run wizard** - Easy setup to select your Bluetooth device
- **Live log viewer** - Monitor connection status in real-time
- **Start on boot** - Optional automatic startup with Windows
- **Low resource usage** - Minimal CPU and memory footprint
- **Single instance** - Prevents multiple instances from running

## How It Works

1. NearLock continuously monitors the connection status of your selected Bluetooth device
2. When the device disconnects (you walk away), a countdown starts (default: 20 seconds)
3. After the countdown, a confirmation scan is performed to avoid false positives
4. If the device is still not detected, your PC is automatically locked
5. When you return and your device reconnects, simply unlock your PC normally

## Requirements

- **Windows 10/11** (64-bit recommended)
- **Bluetooth adapter** enabled on your PC
- **Paired Bluetooth device** (phone, smartwatch, etc.)
- **PowerShell 5.1+** (included with Windows 10/11)

## Installation

### Option 1: Download Release (Recommended)

1. Download `NearLock.exe` from the [Releases](../../releases) page
2. Place it in a folder of your choice (e.g., `C:\Program Files\NearLock`)
3. Run `NearLock.exe`
4. Follow the first-run wizard to select your Bluetooth device

### Option 2: Run from Source

1. Clone this repository
2. Run `NearLock-Single.ps1` with PowerShell:
   ```powershell
   powershell -ExecutionPolicy Bypass -File NearLock-Single.ps1
   ```

## Usage

### System Tray Icon

NearLock runs in the system tray with a colored eye icon indicating status:

| Icon Color | Status |
|------------|--------|
| 🟢 Green | Device connected |
| 🟠 Orange | Searching for device |
| ⚫ Grey | Disabled or no device configured |

### Tray Menu Options

- **Auto-lock: Enabled/Disabled** - Toggle automatic locking on/off
- **Settings...** - Change Bluetooth device or configure startup options
- **View Logs** - Open live log window to monitor activity
- **Exit** - Close NearLock

### Settings

Access settings from the tray menu to:

- **Change Device** - Select a different paired Bluetooth device
- **Start on boot** - Enable/disable automatic startup with Windows

## Building from Source

### Prerequisites

- PowerShell 5.1 or later
- [ps2exe](https://github.com/MScholtes/PS2EXE) (for compiling to EXE)

### Build Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/baptisteba/NearLock.git
   cd NearLock
   ```

2. Install ps2exe (if not already installed):
   ```powershell
   Install-Module ps2exe -Scope CurrentUser
   ```

3. Compile to EXE:
   ```powershell
   Invoke-ps2exe -inputFile "NearLock-Single.ps1" -outputFile "NearLock.exe" -noConsole -title "NearLock" -version "1.2.1.0"
   ```

## Configuration

Configuration is stored in `%LOCALAPPDATA%\NearLock\config.json`:

```json
{
  "deviceMAC": "AA:BB:CC:DD:EE:FF",
  "deviceName": "My Phone",
  "startOnBoot": true
}
```

Logs are stored in `%LOCALAPPDATA%\NearLock\logs\`.

## Technical Details

### Detection Methods

NearLock uses multiple detection methods for reliability:

1. **Bluetooth Classic API** - Uses Windows `bthprops.cpl` to check device connection status
2. **BLE PnP Device** - Monitors Bluetooth Low Energy device status via PnP subsystem
3. **Radio Scan** - Performs active Bluetooth inquiry as fallback confirmation

### Timing Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| Poll interval | 4 seconds | How often to check device status |
| Lock threshold | 20 seconds | Time before locking after disconnect |
| Grace period | 30 seconds | Delay after system wake before monitoring |

### Security

- Uses Windows native `LockWorkStation()` API to lock the PC
- No network communication - all processing is local
- No data collection or telemetry

## Troubleshooting

### Device not detected

1. Ensure your Bluetooth device is paired in Windows Settings
2. Make sure Bluetooth is enabled on both your PC and device
3. Try re-pairing the device

### False locks

If NearLock locks your PC while you're still nearby:
- Your device may have intermittent Bluetooth connectivity
- Try moving closer to your PC
- Check for Bluetooth interference from other devices

### High CPU usage

If you notice high CPU usage:
- Check the logs for repeated errors
- Try restarting NearLock
- Ensure your Bluetooth driver is up to date

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- Inspired by the macOS/iOS Near Lock app
- Uses Windows Bluetooth APIs via P/Invoke
- Built with PowerShell and Windows Forms
