# NearLock - Single EXE version
# All-in-one: Tray + Monitor + Watchdog using background runspace
# Detection: Classic BT + BLE + BLE-Proximity (WinRT GATT)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Runtime.WindowsRuntime

# Load WinRT types for BLE proximity check
$null = [Windows.Devices.Bluetooth.BluetoothLEDevice, Windows.Devices.Bluetooth, ContentType=WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattDeviceServicesResult, Windows.Devices.Bluetooth, ContentType=WindowsRuntime]

# --- Paths ---
$script:dataDir = Join-Path $env:LOCALAPPDATA "NearLock"
if (-not (Test-Path $script:dataDir)) { New-Item -ItemType Directory -Path $script:dataDir -Force | Out-Null }
$script:logDir = Join-Path $script:dataDir "logs"
if (-not (Test-Path $script:logDir)) { New-Item -ItemType Directory -Path $script:logDir -Force | Out-Null }
$script:configPath = Join-Path $script:dataDir "config.json"
$script:autoLockEnabled = $true
$script:monitorRunspace = $null
$script:monitorPowerShell = $null
$script:logWindow = $null

# --- Single instance check ---
$script:mutexName = "Global\NearLock-mutex"
$script:mutex = New-Object System.Threading.Mutex($false, $script:mutexName)
$script:hasMutex = $false
try {
    $script:hasMutex = $script:mutex.WaitOne(0, $false)
} catch [System.Threading.AbandonedMutexException] {
    $script:hasMutex = $true
}
if (-not $script:hasMutex) {
    [System.Windows.Forms.MessageBox]::Show("NearLock is already running.", "NearLock", 0, 64)
    exit
}

# --- Bluetooth API ---
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public static class BT {
    [DllImport("user32.dll")]
    public static extern bool LockWorkStation();

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SEARCH_PARAMS {
        public int dwSize;
        public bool fReturnAuthenticated;
        public bool fReturnRemembered;
        public bool fReturnUnknown;
        public bool fReturnConnected;
        public bool fIssueInquiry;
        public byte cTimeoutMultiplier;
        public IntPtr hRadio;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DEVICE_INFO {
        public int dwSize;
        public long Address;
        public uint ulClassofDevice;
        [MarshalAs(UnmanagedType.Bool)] public bool fConnected;
        [MarshalAs(UnmanagedType.Bool)] public bool fRemembered;
        [MarshalAs(UnmanagedType.Bool)] public bool fAuthenticated;
        public SYSTEMTIME stLastSeen;
        public SYSTEMTIME stLastUsed;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 248)] public string szName;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SYSTEMTIME {
        public ushort wYear, wMonth, wDayOfWeek, wDay, wHour, wMinute, wSecond, wMilliseconds;
    }

    [DllImport("bthprops.cpl", CharSet = CharSet.Unicode)]
    public static extern IntPtr BluetoothFindFirstDevice(ref SEARCH_PARAMS p, ref DEVICE_INFO d);
    [DllImport("bthprops.cpl", CharSet = CharSet.Unicode)]
    public static extern bool BluetoothFindNextDevice(IntPtr h, ref DEVICE_INFO d);
    [DllImport("bthprops.cpl")]
    public static extern bool BluetoothFindDeviceClose(IntPtr h);
    [DllImport("bthprops.cpl", CharSet = CharSet.Unicode)]
    public static extern int BluetoothGetDeviceInfo(IntPtr hRadio, ref DEVICE_INFO d);

    public static long ParseMAC(string mac) {
        return Convert.ToInt64(mac.Replace(":", "").Replace("-", ""), 16);
    }

    public static string FormatMAC(long addr) {
        byte[] b = BitConverter.GetBytes(addr);
        return string.Format("{0:X2}:{1:X2}:{2:X2}:{3:X2}:{4:X2}:{5:X2}", b[5], b[4], b[3], b[2], b[1], b[0]);
    }

    public static bool IsConnected(long addr) {
        if (addr == 0) return false;
        DEVICE_INFO info = new DEVICE_INFO();
        info.dwSize = Marshal.SizeOf(typeof(DEVICE_INFO));
        info.Address = addr;
        return BluetoothGetDeviceInfo(IntPtr.Zero, ref info) == 0 && info.fConnected;
    }

    public static bool IsNearby(long addr, byte timeout) {
        if (addr == 0) return false;
        SEARCH_PARAMS p = new SEARCH_PARAMS();
        p.dwSize = Marshal.SizeOf(typeof(SEARCH_PARAMS));
        p.fReturnUnknown = true;
        p.fReturnConnected = true;
        p.fIssueInquiry = true;
        p.cTimeoutMultiplier = timeout;

        DEVICE_INFO d = new DEVICE_INFO();
        d.dwSize = Marshal.SizeOf(typeof(DEVICE_INFO));

        IntPtr h = BluetoothFindFirstDevice(ref p, ref d);
        if (h != IntPtr.Zero) {
            try {
                do { if (d.Address == addr) return true; } while (BluetoothFindNextDevice(h, ref d));
            } finally { BluetoothFindDeviceClose(h); }
        }
        return IsConnected(addr);
    }

    public static List<Tuple<string, string, bool>> GetPaired() {
        var list = new List<Tuple<string, string, bool>>();
        SEARCH_PARAMS p = new SEARCH_PARAMS();
        p.dwSize = Marshal.SizeOf(typeof(SEARCH_PARAMS));
        p.fReturnAuthenticated = true;
        p.fReturnRemembered = true;
        p.fReturnConnected = true;

        DEVICE_INFO d = new DEVICE_INFO();
        d.dwSize = Marshal.SizeOf(typeof(DEVICE_INFO));

        IntPtr h = BluetoothFindFirstDevice(ref p, ref d);
        if (h != IntPtr.Zero) {
            try {
                do { list.Add(new Tuple<string, string, bool>(d.szName ?? "Unknown", FormatMAC(d.Address), d.fConnected)); }
                while (BluetoothFindNextDevice(h, ref d));
            } finally { BluetoothFindDeviceClose(h); }
        }
        return list;
    }
}
"@ -ErrorAction SilentlyContinue

# --- Helper functions ---
function Write-Log($msg) {
    $logFile = Join-Path $script:logDir "NearLock_$(Get-Date -Format 'yyyy-MM-dd').log"
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    try { Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Get-Config {
    if (Test-Path $script:configPath) {
        try { return Get-Content $script:configPath -Raw | ConvertFrom-Json } catch {}
    }
    return $null
}

function Save-Config {
    param($mac, $name, $bleMac = $null, $startOnBoot = $null)
    $cfg = Get-Config
    $data = @{
        deviceMAC = if ($mac) { $mac } elseif ($cfg) { $cfg.deviceMAC } else { $null }
        deviceName = if ($name) { $name } elseif ($cfg) { $cfg.deviceName } else { $null }
        deviceBLEMAC = if ($bleMac) { $bleMac } elseif ($cfg -and $cfg.deviceBLEMAC) { $cfg.deviceBLEMAC } else { $null }
        startOnBoot = if ($null -ne $startOnBoot) { $startOnBoot } elseif ($cfg -and $null -ne $cfg.startOnBoot) { $cfg.startOnBoot } else { $false }
    }
    $data | ConvertTo-Json | Set-Content $script:configPath -Encoding UTF8
}

# --- Startup management ---
$script:startupRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$script:startupRegName = "NearLock"

function Get-StartOnBoot {
    try {
        $val = Get-ItemProperty -Path $script:startupRegPath -Name $script:startupRegName -ErrorAction SilentlyContinue
        return ($null -ne $val)
    } catch { return $false }
}

function Set-StartOnBoot($enabled) {
    try {
        if ($enabled) {
            $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
            Set-ItemProperty -Path $script:startupRegPath -Name $script:startupRegName -Value "`"$exePath`"" -ErrorAction Stop
        } else {
            Remove-ItemProperty -Path $script:startupRegPath -Name $script:startupRegName -ErrorAction SilentlyContinue
        }
        Save-Config -startOnBoot $enabled
        return $true
    } catch { return $false }
}

function Get-PairedDevices {
    try {
        [BT]::GetPaired() | ForEach-Object { [PSCustomObject]@{ Name = $_.Item1; MAC = $_.Item2; Connected = $_.Item3; Source = "Paired" } }
    } catch { @() }
}

# Helper to await WinRT async operations (for main thread)
$script:asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]

function Await-WinRT($WinRtTask, $ResultType) {
    $asTask = $script:asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(15000) | Out-Null
    if ($netTask.IsCompleted) { return $netTask.Result }
    return $null
}

# Scan for nearby Classic Bluetooth devices via inquiry
function Get-NearbyClassicBT {
    param([byte]$TimeoutMultiplier = 4)
    $devices = @()
    try {
        $searchParams = New-Object BT+SEARCH_PARAMS
        $searchParams.dwSize = [System.Runtime.InteropServices.Marshal]::SizeOf([Type][BT+SEARCH_PARAMS])
        $searchParams.fReturnAuthenticated = $false
        $searchParams.fReturnRemembered = $false
        $searchParams.fReturnUnknown = $true
        $searchParams.fReturnConnected = $true
        $searchParams.fIssueInquiry = $true
        $searchParams.cTimeoutMultiplier = $TimeoutMultiplier

        $deviceInfo = New-Object BT+DEVICE_INFO
        $deviceInfo.dwSize = [System.Runtime.InteropServices.Marshal]::SizeOf([Type][BT+DEVICE_INFO])

        $handle = [BT]::BluetoothFindFirstDevice([ref]$searchParams, [ref]$deviceInfo)
        if ($handle -ne [IntPtr]::Zero) {
            try {
                do {
                    $mac = [BT]::FormatMAC($deviceInfo.Address)
                    $name = if ($deviceInfo.szName) { $deviceInfo.szName } else { "(Unknown)" }
                    $devices += [PSCustomObject]@{
                        Name = $name
                        MAC = $mac
                        Connected = $deviceInfo.fConnected
                        Source = "Classic BT"
                        Address = $deviceInfo.Address
                    }
                } while ([BT]::BluetoothFindNextDevice($handle, [ref]$deviceInfo))
            } finally {
                [BT]::BluetoothFindDeviceClose($handle) | Out-Null
            }
        }
    } catch {}
    return $devices
}

# Scan for nearby BLE devices via WinRT
function Get-NearbyBLE {
    $devices = @()
    try {
        $null = [Windows.Devices.Enumeration.DeviceInformation, Windows.Devices.Enumeration, ContentType=WindowsRuntime]
        $null = [Windows.Devices.Enumeration.DeviceInformationCollection, Windows.Devices.Enumeration, ContentType=WindowsRuntime]

        # Get unpaired BLE devices
        $bleSelector = [Windows.Devices.Bluetooth.BluetoothLEDevice]::GetDeviceSelectorFromPairingState($false)
        $bleDevices = Await-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($bleSelector)) ([Windows.Devices.Enumeration.DeviceInformationCollection])

        if ($bleDevices) {
            foreach ($dev in $bleDevices) {
                # Extract MAC from device ID (format: BluetoothLE#BluetoothLExx:xx:xx:xx:xx:xx-yy:yy:yy:yy:yy:yy)
                if ($dev.Id -match '([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}$') {
                    $mac = $Matches[0] -replace '-', ':'
                    $mac = $mac.ToUpper()
                    $name = if ($dev.Name) { $dev.Name } else { "(Unknown BLE)" }
                    $devices += [PSCustomObject]@{
                        Name = $name
                        MAC = $mac
                        Connected = $false
                        Source = "BLE"
                        DeviceId = $dev.Id
                    }
                }
            }
        }

        # Also get paired BLE devices
        $blePairedSelector = [Windows.Devices.Bluetooth.BluetoothLEDevice]::GetDeviceSelectorFromPairingState($true)
        $blePairedDevices = Await-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($blePairedSelector)) ([Windows.Devices.Enumeration.DeviceInformationCollection])

        if ($blePairedDevices) {
            foreach ($dev in $blePairedDevices) {
                if ($dev.Id -match '([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}$') {
                    $mac = $Matches[0] -replace '-', ':'
                    $mac = $mac.ToUpper()
                    # Check if already in list
                    if (-not ($devices | Where-Object { $_.MAC -eq $mac })) {
                        $name = if ($dev.Name) { $dev.Name } else { "(Unknown BLE)" }
                        $devices += [PSCustomObject]@{
                            Name = $name
                            MAC = $mac
                            Connected = $false
                            Source = "BLE (Paired)"
                            DeviceId = $dev.Id
                        }
                    }
                }
            }
        }
    } catch {}
    return $devices
}

# Merge Classic BT and BLE devices, linking related MACs
function Merge-DeviceLists {
    param($ClassicDevices, $BLEDevices, $PairedDevices)

    $merged = @{}

    # Add paired devices first
    foreach ($d in $PairedDevices) {
        $key = $d.MAC
        $merged[$key] = [PSCustomObject]@{
            Name = $d.Name
            MAC = $d.MAC
            BLEMAC = $null
            Connected = $d.Connected
            Sources = @("Paired")
            IsPaired = $true
        }
    }

    # Add/merge Classic BT devices
    foreach ($d in $ClassicDevices) {
        $key = $d.MAC
        if ($merged.ContainsKey($key)) {
            if ("Classic BT" -notin $merged[$key].Sources) {
                $merged[$key].Sources += "Classic BT"
            }
            if ($d.Connected) { $merged[$key].Connected = $true }
        } else {
            $merged[$key] = [PSCustomObject]@{
                Name = $d.Name
                MAC = $d.MAC
                BLEMAC = $null
                Connected = $d.Connected
                Sources = @("Classic BT")
                IsPaired = $false
            }
        }
    }

    # Add/merge BLE devices - try to link by name or add as separate
    foreach ($d in $BLEDevices) {
        $linked = $false
        # Try to find matching device by name
        foreach ($key in @($merged.Keys)) {
            $existing = $merged[$key]
            if ($existing.Name -eq $d.Name -and $d.Name -ne "(Unknown BLE)" -and $d.Name -ne "(Unknown)") {
                # Link BLE MAC to existing device
                $existing.BLEMAC = $d.MAC
                if ("BLE" -notin $existing.Sources) {
                    $existing.Sources += "BLE"
                }
                $linked = $true
                break
            }
        }

        if (-not $linked) {
            # Check if MAC already exists (some devices use same MAC for BT and BLE)
            if ($merged.ContainsKey($d.MAC)) {
                if ("BLE" -notin $merged[$d.MAC].Sources) {
                    $merged[$d.MAC].Sources += "BLE"
                }
            } else {
                # Add as new BLE-only device
                $merged[$d.MAC] = [PSCustomObject]@{
                    Name = $d.Name
                    MAC = $d.MAC
                    BLEMAC = $d.MAC
                    Connected = $false
                    Sources = @("BLE")
                    IsPaired = ($d.Source -eq "BLE (Paired)")
                }
            }
        }
    }

    return $merged.Values | Sort-Object -Property @{Expression={$_.IsPaired}; Descending=$true}, @{Expression={$_.Connected}; Descending=$true}, Name
}

function Show-DeviceDialog {
    param([switch]$IsWizard)

    $title = if ($IsWizard) { "NearLock Setup - Select Device" } else { "Scan for Nearby Devices" }
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text = $title
        Size = New-Object System.Drawing.Size(500, 420)
        StartPosition = "CenterScreen"
        FormBorderStyle = "FixedDialog"
        MaximizeBox = $false
        MinimizeBox = $false
        TopMost = $true
    }

    # Scan button and status
    $scanBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "Start Scan (~15s)"
        Location = New-Object System.Drawing.Point(10, 10)
        Size = New-Object System.Drawing.Size(120, 28)
    }

    $scanStatus = New-Object System.Windows.Forms.Label -Property @{
        Text = "Click 'Start Scan' to discover nearby devices"
        Location = New-Object System.Drawing.Point(140, 16)
        Size = New-Object System.Drawing.Size(330, 20)
        ForeColor = [System.Drawing.Color]::Gray
    }

    # Device list
    $scanList = New-Object System.Windows.Forms.ListView -Property @{
        Location = New-Object System.Drawing.Point(10, 45)
        Size = New-Object System.Drawing.Size(465, 280)
        View = "Details"
        FullRowSelect = $true
        GridLines = $true
    }
    $scanList.Columns.Add("Name", 160) | Out-Null
    $scanList.Columns.Add("MAC Address", 130) | Out-Null
    $scanList.Columns.Add("BLE MAC", 110) | Out-Null
    $scanList.Columns.Add("Sources", 90) | Out-Null

    # Buttons
    $okBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "OK"
        Location = New-Object System.Drawing.Point(300, 340)
        Size = New-Object System.Drawing.Size(80, 30)
    }

    $cancelBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "Cancel"
        Location = New-Object System.Drawing.Point(390, 340)
        Size = New-Object System.Drawing.Size(80, 30)
        DialogResult = "Cancel"
    }

    $form.Controls.AddRange(@($scanBtn, $scanStatus, $scanList, $okBtn, $cancelBtn))
    $form.CancelButton = $cancelBtn

    # Store selected device and scanned devices
    $script:selectedDevice = $null
    $script:scannedDevices = @()

    # Get paired devices for merging
    $pairedDevices = @(Get-PairedDevices)

    # Scan button click
    $scanBtn.Add_Click({
        $scanBtn.Enabled = $false
        $scanBtn.Text = "Scanning..."
        $scanStatus.Text = "Scanning for Classic Bluetooth devices..."
        $scanStatus.ForeColor = [System.Drawing.Color]::Blue
        $scanList.Items.Clear()
        $form.Refresh()

        # Scan Classic BT (takes ~5-10s)
        $classicDevices = @(Get-NearbyClassicBT -TimeoutMultiplier 4)

        $scanStatus.Text = "Scanning for BLE devices..."
        $form.Refresh()

        # Scan BLE
        $bleDevices = @(Get-NearbyBLE)

        $scanStatus.Text = "Merging results..."
        $form.Refresh()

        # Merge all devices
        $script:scannedDevices = @(Merge-DeviceLists -ClassicDevices $classicDevices -BLEDevices $bleDevices -PairedDevices $pairedDevices)

        # Populate list
        $cfg = Get-Config
        foreach ($d in $script:scannedDevices) {
            $item = New-Object System.Windows.Forms.ListViewItem($d.Name)
            $item.SubItems.Add($d.MAC) | Out-Null
            $item.SubItems.Add($(if ($d.BLEMAC -and $d.BLEMAC -ne $d.MAC) { $d.BLEMAC } else { "-" })) | Out-Null
            $item.SubItems.Add(($d.Sources -join ", ")) | Out-Null
            $item.Tag = $d
            if ($d.Connected) { $item.ForeColor = [System.Drawing.Color]::Green }
            elseif ($d.IsPaired) { $item.ForeColor = [System.Drawing.Color]::Blue }
            # Pre-select current device if configured
            if ($cfg -and ($d.MAC -eq $cfg.deviceMAC -or $d.BLEMAC -eq $cfg.deviceMAC)) { $item.Selected = $true }
            $scanList.Items.Add($item) | Out-Null
        }

        $count = $script:scannedDevices.Count
        $scanStatus.Text = "Found $count device(s). Select one and click OK."
        $scanStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        $scanBtn.Text = "Rescan"
        $scanBtn.Enabled = $true
    })

    # OK button click
    $okBtn.Add_Click({
        if ($scanList.SelectedItems.Count -gt 0) {
            $selected = $scanList.SelectedItems[0].Tag
            # Use BLE MAC if available and different (for BLE-only proximity)
            $macToUse = $selected.MAC
            if ($selected.BLEMAC -and $selected.Sources -contains "BLE") {
                $macToUse = if ($selected.BLEMAC) { $selected.BLEMAC } else { $selected.MAC }
            }
            $script:selectedDevice = @{
                Name = $selected.Name
                MAC = $selected.MAC
                BLEMAC = $selected.BLEMAC
                Sources = $selected.Sources
            }
            $form.DialogResult = "OK"
            $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please scan and select a device.", "NearLock", 0, 64)
        }
    })

    # Double-click to select
    $scanList.Add_DoubleClick({ $okBtn.PerformClick() })

    if ($form.ShowDialog() -eq "OK" -and $script:selectedDevice) {
        return $script:selectedDevice
    }
    return $null
}

# --- Settings Dialog ---
function Show-SettingsDialog {
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text = "NearLock Settings"
        Size = New-Object System.Drawing.Size(420, 320)
        StartPosition = "CenterScreen"
        FormBorderStyle = "FixedDialog"
        MaximizeBox = $false
        MinimizeBox = $false
        TopMost = $true
    }

    $cfg = Get-Config
    $script:settingsChanged = $false
    $script:deviceChanged = $false

    # --- Device section ---
    $deviceGroup = New-Object System.Windows.Forms.GroupBox -Property @{
        Text = "Bluetooth Device"
        Location = New-Object System.Drawing.Point(15, 15)
        Size = New-Object System.Drawing.Size(375, 120)
    }

    $deviceLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "Device:"
        Location = New-Object System.Drawing.Point(10, 22)
        Size = New-Object System.Drawing.Size(50, 20)
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
    }

    $currentDevice = if ($cfg -and $cfg.deviceName) { "$($cfg.deviceName)" } else { "(No device selected)" }
    $script:deviceValue = New-Object System.Windows.Forms.Label -Property @{
        Text = $currentDevice
        Location = New-Object System.Drawing.Point(65, 22)
        Size = New-Object System.Drawing.Size(295, 20)
        Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    }

    $macLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "MAC:"
        Location = New-Object System.Drawing.Point(10, 44)
        Size = New-Object System.Drawing.Size(50, 18)
        Font = New-Object System.Drawing.Font("Segoe UI", 8)
        ForeColor = [System.Drawing.Color]::Gray
    }

    $currentMAC = if ($cfg -and $cfg.deviceMAC) { $cfg.deviceMAC } else { "-" }
    $script:macValue = New-Object System.Windows.Forms.Label -Property @{
        Text = $currentMAC
        Location = New-Object System.Drawing.Point(65, 44)
        Size = New-Object System.Drawing.Size(295, 18)
        Font = New-Object System.Drawing.Font("Consolas", 8)
        ForeColor = [System.Drawing.Color]::Gray
    }

    $bleMacLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "BLE:"
        Location = New-Object System.Drawing.Point(10, 62)
        Size = New-Object System.Drawing.Size(50, 18)
        Font = New-Object System.Drawing.Font("Segoe UI", 8)
        ForeColor = [System.Drawing.Color]::Gray
    }

    $currentBLEMAC = if ($cfg -and $cfg.deviceBLEMAC) { $cfg.deviceBLEMAC } else { "-" }
    $script:bleMacValue = New-Object System.Windows.Forms.Label -Property @{
        Text = $currentBLEMAC
        Location = New-Object System.Drawing.Point(65, 62)
        Size = New-Object System.Drawing.Size(295, 18)
        Font = New-Object System.Drawing.Font("Consolas", 8)
        ForeColor = [System.Drawing.Color]::Gray
    }

    $changeBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "Change Device..."
        Location = New-Object System.Drawing.Point(10, 85)
        Size = New-Object System.Drawing.Size(120, 25)
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
    }
    $changeBtn.Add_Click({
        $sel = Show-DeviceDialog
        if ($sel) {
            Save-Config -mac $sel.MAC -name $sel.Name -bleMac $sel.BLEMAC
            $script:deviceValue.Text = $sel.Name
            $script:macValue.Text = if ($sel.MAC) { $sel.MAC } else { "-" }
            $script:bleMacValue.Text = if ($sel.BLEMAC) { $sel.BLEMAC } else { "-" }
            $script:deviceChanged = $true
            $script:settingsChanged = $true
        }
    })

    $deviceGroup.Controls.AddRange(@($deviceLabel, $script:deviceValue, $macLabel, $script:macValue, $bleMacLabel, $script:bleMacValue, $changeBtn))

    # --- Startup section ---
    $startupGroup = New-Object System.Windows.Forms.GroupBox -Property @{
        Text = "Startup"
        Location = New-Object System.Drawing.Point(15, 145)
        Size = New-Object System.Drawing.Size(375, 70)
    }

    $startupCheck = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Start NearLock automatically when Windows starts"
        Location = New-Object System.Drawing.Point(10, 28)
        Size = New-Object System.Drawing.Size(350, 24)
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
        Checked = (Get-StartOnBoot)
    }
    $startupCheck.Add_CheckedChanged({ $script:settingsChanged = $true })

    $startupGroup.Controls.Add($startupCheck)

    # --- Buttons ---
    $okBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "OK"
        Location = New-Object System.Drawing.Point(220, 235)
        Size = New-Object System.Drawing.Size(80, 30)
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
    }
    $okBtn.Add_Click({
        Set-StartOnBoot $startupCheck.Checked
        $form.DialogResult = "OK"
        $form.Close()
    })

    $cancelBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "Cancel"
        Location = New-Object System.Drawing.Point(310, 235)
        Size = New-Object System.Drawing.Size(80, 30)
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
    }
    $cancelBtn.Add_Click({
        $form.DialogResult = "Cancel"
        $form.Close()
    })

    $form.Controls.AddRange(@($deviceGroup, $startupGroup, $okBtn, $cancelBtn))
    $form.AcceptButton = $okBtn
    $form.CancelButton = $cancelBtn

    $result = $form.ShowDialog()
    return @{ Result = $result; DeviceChanged = $script:deviceChanged }
}

# --- First Run Wizard ---
function Show-FirstRunWizard {
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text = "Welcome to NearLock"
        Size = New-Object System.Drawing.Size(450, 320)
        StartPosition = "CenterScreen"
        FormBorderStyle = "FixedDialog"
        MaximizeBox = $false
        MinimizeBox = $false
        TopMost = $true
    }

    # Icon/Logo area
    $iconPanel = New-Object System.Windows.Forms.Panel -Property @{
        Location = New-Object System.Drawing.Point(20, 20)
        Size = New-Object System.Drawing.Size(60, 60)
        BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    }
    $iconLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = [char]0x1F441  # Eye emoji (will show as box but that's ok)
        Font = New-Object System.Drawing.Font("Segoe UI", 24)
        ForeColor = [System.Drawing.Color]::White
        TextAlign = "MiddleCenter"
        Dock = "Fill"
    }
    $iconPanel.Controls.Add($iconLabel)

    $titleLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "Welcome to NearLock"
        Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
        Location = New-Object System.Drawing.Point(95, 25)
        Size = New-Object System.Drawing.Size(320, 30)
    }

    $subtitleLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "Automatic PC lock when you walk away"
        Font = New-Object System.Drawing.Font("Segoe UI", 10)
        ForeColor = [System.Drawing.Color]::Gray
        Location = New-Object System.Drawing.Point(95, 55)
        Size = New-Object System.Drawing.Size(320, 20)
    }

    $descLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "NearLock monitors your Bluetooth device (like your phone) and automatically locks your PC when you walk away.`n`nTo get started, select the Bluetooth device you want to use for presence detection. Make sure the device is already paired with your PC."
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
        Location = New-Object System.Drawing.Point(20, 100)
        Size = New-Object System.Drawing.Size(400, 80)
    }

    $noteLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "Tip: Your phone works best as it's usually always with you."
        Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
        ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        Location = New-Object System.Drawing.Point(20, 185)
        Size = New-Object System.Drawing.Size(400, 20)
    }

    $setupBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "Select Device..."
        Font = New-Object System.Drawing.Font("Segoe UI", 10)
        Location = New-Object System.Drawing.Point(20, 225)
        Size = New-Object System.Drawing.Size(130, 35)
        BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        ForeColor = [System.Drawing.Color]::White
        FlatStyle = "Flat"
    }

    $skipBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "Skip for now"
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
        Location = New-Object System.Drawing.Point(165, 225)
        Size = New-Object System.Drawing.Size(100, 35)
        FlatStyle = "Flat"
    }

    $exitBtn = New-Object System.Windows.Forms.Button -Property @{
        Text = "Exit"
        Font = New-Object System.Drawing.Font("Segoe UI", 9)
        Location = New-Object System.Drawing.Point(340, 225)
        Size = New-Object System.Drawing.Size(80, 35)
        FlatStyle = "Flat"
    }

    $script:wizardResult = "skip"

    $setupBtn.Add_Click({
        $form.Hide()
        $sel = Show-DeviceDialog -IsWizard
        if ($sel) {
            Save-Config -mac $sel.MAC -name $sel.Name -bleMac $sel.BLEMAC
            $script:wizardResult = "configured"
            $form.Close()
        } else {
            $form.Show()
        }
    })

    $skipBtn.Add_Click({
        $script:wizardResult = "skip"
        $form.Close()
    })

    $exitBtn.Add_Click({
        $script:wizardResult = "exit"
        $form.Close()
    })

    $form.Controls.AddRange(@($iconPanel, $titleLabel, $subtitleLabel, $descLabel, $noteLabel, $setupBtn, $skipBtn, $exitBtn))
    $form.ShowDialog() | Out-Null

    return $script:wizardResult
}

function Test-NeedsWizard {
    $cfg = Get-Config
    if (-not $cfg) { return $true }
    if (-not $cfg.deviceMAC) { return $true }
    if ($cfg.deviceMAC -eq "00:00:00:00:00:00") { return $true }
    return $false
}

# --- Live Log Window ---
function Show-LogWindow {
    if ($script:logWindow -and -not $script:logWindow.IsDisposed) {
        $script:logWindow.Activate()
        return
    }

    $script:logWindow = New-Object System.Windows.Forms.Form -Property @{
        Text = "NearLock - Live Logs"
        Size = New-Object System.Drawing.Size(600, 400)
        StartPosition = "CenterScreen"
        FormBorderStyle = "Sizable"
        MinimumSize = New-Object System.Drawing.Size(400, 300)
    }

    $textBox = New-Object System.Windows.Forms.TextBox -Property @{
        Multiline = $true
        ReadOnly = $true
        ScrollBars = "Vertical"
        Dock = "Fill"
        Font = New-Object System.Drawing.Font("Consolas", 9)
        BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    }

    $script:logLastSize = 0
    $script:logTextBox = $textBox

    # Initial load
    $logFile = Join-Path $script:logDir "NearLock_$(Get-Date -Format 'yyyy-MM-dd').log"
    if (Test-Path $logFile) {
        $content = Get-Content $logFile -Raw -ErrorAction SilentlyContinue
        $textBox.Text = $content
        $script:logLastSize = (Get-Item $logFile -ErrorAction SilentlyContinue).Length
        $textBox.SelectionStart = $textBox.Text.Length
        $textBox.ScrollToCaret()
    }

    # Timer to refresh
    $script:logTimer = New-Object System.Windows.Forms.Timer -Property @{ Interval = 1000 }
    $script:logTimer.Add_Tick({
        $logFile = Join-Path $script:logDir "NearLock_$(Get-Date -Format 'yyyy-MM-dd').log"
        if (Test-Path $logFile) {
            $currentSize = (Get-Item $logFile -ErrorAction SilentlyContinue).Length
            if ($currentSize -ne $script:logLastSize) {
                $content = Get-Content $logFile -Raw -ErrorAction SilentlyContinue
                $script:logTextBox.Text = $content
                $script:logLastSize = $currentSize
                $script:logTextBox.SelectionStart = $script:logTextBox.Text.Length
                $script:logTextBox.ScrollToCaret()
            }
        }
    })
    $script:logTimer.Start()

    $script:logWindow.Add_FormClosing({
        try { $script:logTimer.Stop(); $script:logTimer.Dispose() } catch {}
        $script:logTimer = $null
    })
    $script:logWindow.Controls.Add($textBox)
    $script:logWindow.Show()
}

# --- Monitor script block (runs in background) ---
$script:monitorScript = {
    param($configPath, $logDir)

    # Load WinRT for BLE proximity in runspace
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Devices.Bluetooth.BluetoothLEDevice, Windows.Devices.Bluetooth, ContentType=WindowsRuntime]
    $null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattDeviceServicesResult, Windows.Devices.Bluetooth, ContentType=WindowsRuntime]

    # Helper to await WinRT async operations
    $script:asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]

    function Await-WinRT($WinRtTask, $ResultType) {
        $asTask = $script:asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(10000) | Out-Null
        if ($netTask.IsCompleted) { return $netTask.Result }
        return $null
    }

    function Write-Log($msg) {
        $logFile = Join-Path $logDir "NearLock_$(Get-Date -Format 'yyyy-MM-dd').log"
        try { Add-Content -Path $logFile -Value "[$(Get-Date -Format 'HH:mm:ss')] $msg" -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
    }

    # Load config
    $deviceName = $null
    $targetMAC = $null
    $targetBLEMAC = $null
    if (Test-Path $configPath) {
        try {
            $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
            if ($cfg.deviceMAC -and $cfg.deviceMAC -ne "00:00:00:00:00:00") {
                $targetMAC = $cfg.deviceMAC
                $deviceName = $cfg.deviceName
                $targetBLEMAC = $cfg.deviceBLEMAC  # May be null
            }
        } catch {}
    }

    # No valid device configured - wait for configuration
    if (-not $targetMAC) {
        Write-Log "No device configured - please select a device from the tray menu"
        while ($true) {
            Start-Sleep 5
            if (Test-Path $configPath) {
                try {
                    $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
                    if ($cfg.deviceMAC -and $cfg.deviceMAC -ne "00:00:00:00:00:00") {
                        $targetMAC = $cfg.deviceMAC
                        $deviceName = $cfg.deviceName
                        $targetBLEMAC = $cfg.deviceBLEMAC
                        break
                    }
                } catch {}
            }
        }
    }

    $targetAddr = [BT]::ParseMAC($targetMAC)
    $targetBLE = "DEV_" + ($targetMAC -replace ':', '').ToUpper()

    # Use dedicated BLE MAC if available, otherwise use main MAC
    $bleMacToUse = if ($targetBLEMAC) { $targetBLEMAC } else { $targetMAC }
    $targetBLEAddr = [Convert]::ToUInt64(($bleMacToUse -replace ':', ''), 16)

    # Also prepare secondary BLE ID for PnpDevice check
    $targetBLE2 = if ($targetBLEMAC -and $targetBLEMAC -ne $targetMAC) {
        "DEV_" + ($targetBLEMAC -replace ':', '').ToUpper()
    } else { $null }

    Write-Log "=== NearLock Monitor ==="
    Write-Log "Device: $deviceName ($targetMAC)"
    if ($targetBLEMAC -and $targetBLEMAC -ne $targetMAC) {
        Write-Log "BLE MAC: $targetBLEMAC"
    }
    Write-Log "Detection: Classic BT + BLE + BLE-Proximity"

    $pollInterval = 4; $lockThreshold = 20; $graceAfterResume = 30
    $disconnectedSince = $null; $wasConnected = $false; $everConnected = $false
    $lastPoll = Get-Date; $errors = 0; $startupShown = $false

    # BLE Proximity check via WinRT GATT - detects nearby devices even when not connected
    function Test-BLEProximity {
        try {
            $bleDevice = Await-WinRT ([Windows.Devices.Bluetooth.BluetoothLEDevice]::FromBluetoothAddressAsync($targetBLEAddr)) ([Windows.Devices.Bluetooth.BluetoothLEDevice])
            if ($null -eq $bleDevice) { return $false }
            $servicesResult = Await-WinRT ($bleDevice.GetGattServicesAsync()) ([Windows.Devices.Bluetooth.GenericAttributeProfile.GattDeviceServicesResult])
            if ($null -ne $servicesResult -and $servicesResult.Status -eq [Windows.Devices.Bluetooth.GenericAttributeProfile.GattCommunicationStatus]::Success) {
                return $true
            }
            return $false
        } catch { return $false }
    }

    function Test-Present {
        $classic = [BT]::IsConnected($targetAddr)
        $ble = $false
        try {
            # Check both main MAC and dedicated BLE MAC if different
            $blePatterns = @($targetBLE)
            if ($targetBLE2) { $blePatterns += $targetBLE2 }

            foreach ($pattern in $blePatterns) {
                Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue | Where-Object { $_.InstanceId -match "BTHLE.*$pattern" } | ForEach-Object {
                    $st = Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_DevNodeStatus' -ErrorAction SilentlyContinue
                    if ($st.Data -and (([int]$st.Data -band 8) -ne 0) -and (([int]$st.Data -band 0x400) -eq 0)) {
                        $lc = Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Bluetooth_LastConnectedTime' -ErrorAction SilentlyContinue
                        if ($lc.Data -and ((Get-Date) - [DateTime]$lc.Data).TotalSeconds -lt 30) { $ble = $true }
                    }
                }
                if ($ble) { break }
            }
        } catch {}

        # Fallback: BLE proximity via WinRT GATT (device nearby but not actively connected)
        $bleProximity = $false
        if (-not $classic -and -not $ble) {
            $bleProximity = Test-BLEProximity
        }

        return @{ Classic = $classic; BLE = $ble; BLEProximity = $bleProximity; Connected = ($classic -or $ble -or $bleProximity) }
    }

    while ($true) {
        try {
            $now = Get-Date
            if (($now - $lastPoll).TotalSeconds -gt 15) {
                Write-Log "Wake detected - ${graceAfterResume}s grace"
                $disconnectedSince = $null; $wasConnected = $false; $everConnected = $false; $startupShown = $false
                Start-Sleep $graceAfterResume
                $lastPoll = Get-Date
                continue
            }
            $lastPoll = $now

            $st = Test-Present
            $errors = 0

            if ($st.Connected) {
                if (-not $wasConnected) {
                    $src = @(); if ($st.Classic) { $src += "classic" }; if ($st.BLE) { $src += "BLE" }; if ($st.BLEProximity) { $src += "BLE-proximite" }
                    Write-Log "CONNECTE ($($src -join '+'))"
                }
                $disconnectedSince = $null; $wasConnected = $true; $everConnected = $true
            } else {
                if (-not $everConnected) {
                    if (-not $startupShown) { Write-Log "$deviceName non connecte - scan de recherche actif"; $startupShown = $true }
                    $nearby = [BT]::IsNearby($targetAddr, 2)
                    if (-not $nearby) {
                        # Fallback to BLE GATT proximity
                        $nearby = Test-BLEProximity
                        if ($nearby) {
                            Write-Log "Found via BLE GATT - CONNECTE"
                            $wasConnected = $true; $everConnected = $true; $disconnectedSince = $null
                        }
                    } else {
                        Write-Log "Found via scan - CONNECTE"
                        $wasConnected = $true; $everConnected = $true; $disconnectedSince = $null
                    }
                } else {
                    if ($wasConnected -and $null -eq $disconnectedSince) {
                        $disconnectedSince = $now
                        Write-Log "DECONNECTE - countdown"
                    }
                    if ($null -ne $disconnectedSince) {
                        $elapsed = ($now - $disconnectedSince).TotalSeconds
                        Write-Log "Away $([int]$elapsed)s / ${lockThreshold}s"
                        if ($elapsed -ge $lockThreshold) {
                            Write-Log "Confirmation scan..."
                            $nearby = [BT]::IsNearby($targetAddr, 4)
                            $bleNearby = $false
                            if (-not $nearby) {
                                Write-Log "BLE GATT confirmation..."
                                $bleNearby = Test-BLEProximity
                            }
                            if ($nearby -or $bleNearby) {
                                $method = if ($nearby) { "scan" } else { "BLE GATT" }
                                Write-Log "Found via $method - false alarm"
                                $disconnectedSince = $null; $wasConnected = $true
                            } else {
                                Write-Log "Confirmed - LOCKING"
                                [BT]::LockWorkStation() | Out-Null
                                Start-Sleep 30
                                $lastPoll = Get-Date
                                $disconnectedSince = $null; $wasConnected = $false; $everConnected = $false; $startupShown = $false
                            }
                        }
                    }
                }
                $wasConnected = $false
            }
        } catch {
            $errors++
            Write-Log "Error ($errors): $($_.Exception.Message)"
            if ($errors -ge 3) {
                Write-Log "Too many errors - reset"
                $disconnectedSince = $null; $wasConnected = $false; $everConnected = $false; $startupShown = $false; $errors = 0
                Start-Sleep 10
            }
        }
        Start-Sleep $pollInterval
    }
}

# --- Start/Stop monitor ---
function Start-Monitor {
    if ($script:monitorRunspace) { Stop-Monitor }

    $cfg = Get-Config
    $deviceInfo = if ($cfg.deviceName) { " (device: $($cfg.deviceName))" } else { "" }
    Write-Log "Starting NearLock$deviceInfo"

    $script:monitorRunspace = [runspacefactory]::CreateRunspace()
    $script:monitorRunspace.ApartmentState = "STA"
    $script:monitorRunspace.Open()

    $script:monitorPowerShell = [powershell]::Create()
    $script:monitorPowerShell.Runspace = $script:monitorRunspace
    $script:monitorPowerShell.AddScript($script:monitorScript).AddArgument($script:configPath).AddArgument($script:logDir) | Out-Null
    $script:monitorPowerShell.BeginInvoke() | Out-Null
}

function Stop-Monitor {
    if ($script:monitorPowerShell) {
        try { $script:monitorPowerShell.Stop() } catch {}
        try { $script:monitorPowerShell.Dispose() } catch {}
        $script:monitorPowerShell = $null
    }
    if ($script:monitorRunspace) {
        try { $script:monitorRunspace.Close() } catch {}
        try { $script:monitorRunspace.Dispose() } catch {}
        $script:monitorRunspace = $null
    }
}

# --- Icon generation ---
function New-Icon([System.Drawing.Color]$c) {
    $bmp = New-Object System.Drawing.Bitmap(16,16)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $brush = New-Object System.Drawing.SolidBrush($c)
    $pen = New-Object System.Drawing.Pen($c, 1.5)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0,3,16,14,200,140); $path.AddArc(0,-3,16,14,20,140); $path.CloseFigure()
    $g.DrawPath($pen, $path)
    $g.FillEllipse($brush, 4,4,8,8)
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60,0,0,0))), 6,6,4,4)
    $g.FillEllipse((New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(180,255,255,255))), 5,5,2,2)
    $g.Dispose(); $brush.Dispose(); $pen.Dispose(); $path.Dispose()
    [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

$script:iconGreen = New-Icon ([System.Drawing.Color]::FromArgb(0,180,0))
$script:iconOrange = New-Icon ([System.Drawing.Color]::FromArgb(230,150,0))
$script:iconGrey = New-Icon ([System.Drawing.Color]::FromArgb(140,140,140))

# --- Tray setup ---
$script:tray = New-Object System.Windows.Forms.NotifyIcon -Property @{ Icon = $script:iconGrey; Text = "NearLock"; Visible = $true }
$cms = New-Object System.Windows.Forms.ContextMenuStrip

$script:toggleItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "Auto-lock: Enabled" }
$script:toggleItem.Add_Click({
    if ($script:autoLockEnabled) {
        Stop-Monitor
        $script:autoLockEnabled = $false
        $script:toggleItem.Text = "Auto-lock: Disabled"
        $script:tray.Icon = $script:iconGrey
        $script:tray.Text = "NearLock (disabled)"
    } else {
        Start-Monitor
        $script:autoLockEnabled = $true
        $script:toggleItem.Text = "Auto-lock: Enabled"
        $script:tray.Text = "NearLock"
    }
})
$cms.Items.Add($script:toggleItem) | Out-Null

$settingsItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "Settings..." }
$settingsItem.Add_Click({
    $result = Show-SettingsDialog
    if ($result.DeviceChanged) {
        Stop-Monitor
        Start-Sleep -Milliseconds 300
        Start-Monitor
        $script:autoLockEnabled = $true
        $script:toggleItem.Text = "Auto-lock: Enabled"
        $script:tray.Icon = $script:iconOrange
        $cfg = Get-Config
        if ($cfg) { $script:tray.Text = "NearLock: $($cfg.deviceName)" }
    }
})
$cms.Items.Add($settingsItem) | Out-Null

$logsItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "View Logs" }
$logsItem.Add_Click({ Show-LogWindow })
$cms.Items.Add($logsItem) | Out-Null

$cms.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = "Exit" }
$exitItem.Add_Click({
    try { Stop-Monitor } catch {}
    try { if ($script:logWindow -and -not $script:logWindow.IsDisposed) { $script:logWindow.Close() } } catch {}
    try { $script:tray.Visible = $false; $script:tray.Dispose() } catch {}
    try { if ($script:hasMutex) { $script:mutex.ReleaseMutex() }; $script:mutex.Dispose() } catch {}
    [System.Windows.Forms.Application]::Exit()
})
$cms.Items.Add($exitItem) | Out-Null

$script:tray.ContextMenuStrip = $cms

# --- Status timer ---
$timer = New-Object System.Windows.Forms.Timer -Property @{ Interval = 4000 }
$timer.Add_Tick({
    if (-not $script:autoLockEnabled) { return }
    $f = Join-Path $script:logDir "NearLock_$(Get-Date -Format 'yyyy-MM-dd').log"
    if (-not (Test-Path $f)) { $script:tray.Icon = $script:iconGrey; $script:tray.Text = "NearLock (no log)"; return }
    try {
        $lines = @(Get-Content $f -Tail 10 -ErrorAction SilentlyContinue) | Where-Object { $_ -match '^\[' }
        $last = ($lines | Select-Object -Last 3) -join "`n"
        if ($last -match 'No device configured') { $script:tray.Icon = $script:iconGrey; $script:tray.Text = "NearLock: No device" }
        elseif ($last -match 'non connecte|scan de recherche') { $script:tray.Icon = $script:iconOrange; $script:tray.Text = "NearLock: Searching..." }
        elseif ($last -match 'CONNECTE' -and $last -notmatch 'DECONNECTE') {
            $script:tray.Icon = $script:iconGreen
            if ($last -match 'BLE-proximite|BLE GATT') { $script:tray.Text = "NearLock: Nearby (BLE)" }
            else { $script:tray.Text = "NearLock: Connected" }
        }
        elseif ($last -match 'Starting|Demarrage') { $script:tray.Icon = $script:iconOrange; $script:tray.Text = "NearLock: Starting..." }
        else { $script:tray.Icon = $script:iconGrey; $script:tray.Text = "NearLock" }
    } catch { $script:tray.Icon = $script:iconGrey; $script:tray.Text = "NearLock (error)" }
})
$timer.Start()

# --- First run wizard ---
if (Test-NeedsWizard) {
    $result = Show-FirstRunWizard
    if ($result -eq "exit") {
        try { $timer.Stop(); $timer.Dispose() } catch {}
        try { $script:tray.Visible = $false; $script:tray.Dispose() } catch {}
        try { if ($script:hasMutex) { $script:mutex.ReleaseMutex() }; $script:mutex.Dispose() } catch {}
        exit
    }
    if ($result -eq "configured") {
        $cfg = Get-Config
        if ($cfg) { $script:tray.Text = "NearLock: $($cfg.deviceName)" }
    }
}

# --- Start ---
Start-Monitor
[System.Windows.Forms.Application]::Run()
