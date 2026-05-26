Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- 1. PROSES PENGUMPULAN DATA HARDWARE (ANTI-ERROR) ---
$CompName = $env:COMPUTERNAME

try {
    $sys = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $Brand = "$($sys.Manufacturer) - $($sys.Model)"
} catch { $Brand = "Tidak Terdeteksi" }

try { $OS = (Get-CimInstance Win32_OperatingSystem).Caption } catch { $OS = "Tidak Terdeteksi" }
try { $CPU = (Get-CimInstance Win32_Processor).Name.Trim() } catch { $CPU = "Tidak Terdeteksi" }

try {
    $ramBytes = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum
    $RAM = "$([math]::Round($ramBytes / 1GB)) GB"
} catch { $RAM = "Tidak Terdeteksi" }

try {
    $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
    $diskList = @()
    foreach ($d in $disks) {
        $size = [math]::Round($d.Size / 1GB)
        $free = [math]::Round($d.FreeSpace / 1GB)
        $diskList += "$($d.DeviceID) ($size GB Total, $free GB Free)"
    }
    $Storage = $diskList -join " | "
} catch { $Storage = "Tidak Terdeteksi" }

try {
    $vc = Get-CimInstance Win32_VideoController | Select-Object -First 1
    $gpuName = $vc.Name
    $res = "$($vc.CurrentHorizontalResolution)x$($vc.CurrentVerticalResolution)"
    
    # Mencoba kalkulasi ukuran fisik layar (inch) dari EDID monitor laptop
    $inchStr = ""
    $monitor = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction SilentlyContinue
    if ($monitor) {
        $w = $monitor.MaxHorizontalImageSize
        $h = $monitor.MaxVerticalImageSize
        if ($w -and $h) {
            $inches = [math]::Round([math]::Sqrt([math]::Pow($w, 2) + [math]::Pow($h, 2)) / 2.54, 1)
            if ($inches -gt 0) { $inchStr = " (~$inches inch)" }
        }
    }
    $Layar = "$gpuName [$res]$inchStr"
} catch { $Layar = "Tidak Terdeteksi" }

try {
    $SN = (Get-CimInstance Win32_Bios).SerialNumber
    if ($SN -eq "To be filled by O.E.M." -or [string]::IsNullOrWhiteSpace($SN)) {
        $SN = (Get-CimInstance Win32_ComputerSystemProduct).IdentifyingNumber
    }
} catch { $SN = "Tidak Terdeteksi" }

# --- 2. PEMBUATAN INTERFACE GUI (DARK MODE) ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "System Info - Stock Opname Tool v2.0"
$form.Size = New-Object System.Drawing.Size(620, 580)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Font Settings
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$fontText = New-Object System.Drawing.Font("Segoe UI", 10)
$fontTitle = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)

# Header Title
$title = New-Object System.Windows.Forms.Label
$title.Text = "💻 LAPTOP SPECIFICATION FOR STOCK OPNAME"
$title.Font = $fontTitle
$title.ForeColor = [System.Drawing.Color]::FromArgb(0, 173, 181) # Teal Accent
$title.Size = New-Object System.Drawing.Size(560, 40)
$title.Location = New-Object System.Drawing.Point(20, 20)
$title.TextAlign = "Center"
$form.Controls.Add($title)

# Array data untuk looping UI
$fields = @(
    @("Computer Name", $CompName),
    @("Merk / Tipe", $Brand),
    @("OS Version", $OS),
    @("Processor", $CPU),
    @("RAM Capacity", $RAM),
    @("Storage Info", $Storage),
    @("Display / Grafis", $Layar),
    @("Serial Number", $SN)
)

$y = 80
foreach ($f in $fields) {
    # Label
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $f[0]
    $lbl.Font = $fontLabel
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.Location = New-Object System.Drawing.Point(30, $y)
    $lbl.Size = New-Object System.Drawing.Size(130, 25)
    $form.Controls.Add($lbl)

    # TextBox (Read Only)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text = $f[1]
    $txt.Font = $fontText
    $txt.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $txt.ForeColor = [System.Drawing.Color]::White
    $txt.BorderStyle = "FixedSingle"
    $txt.Location = New-Object System.Drawing.Point(170, $y - 2)
    $txt.Size = New-Object System.Drawing.Size(400, 25)
    $txt.ReadOnly = $true
    $form.Controls.Add($txt)

    $y += 45
}

# Tombol Copy All Data
$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "📋 Copy All Data"
$btnCopy.Font = $fontLabel
$btnCopy.BackColor = [System.Drawing.Color]::FromArgb(0, 173, 181)
$btnCopy.ForeColor = [System.Drawing.Color]::White
$btnCopy.FlatStyle = "Flat"
$btnCopy.Size = New-Object System.Drawing.Size(160, 40)
$btnCopy.Location = New-Object System.Drawing.Point(140, $y + 15)
$btnCopy.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCopy.FlatAppearance.BorderSize = 0
$btnCopy.Add_Click({
    $textToCopy = ""
    foreach ($f in $fields) { $textToCopy += "$($f[0]): $($f[1])`r`n" }
    [System.Windows.Forms.Clipboard]::SetText($textToCopy)
    [System.Windows.Forms.MessageBox]::Show("Semua data spesifikasi berhasil disalin!", "Sukses", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})
$form.Controls.Add($btnCopy)

# Tombol Close
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "❌ Close"
$btnClose.Font = $fontLabel
$btnClose.BackColor = [System.Drawing.Color]::FromArgb(217, 83, 79)
$btnClose.ForeColor = [System.Drawing.Color]::White
$btnClose.FlatStyle = "Flat"
$btnClose.Size = New-Object System.Drawing.Size(160, 40)
$btnClose.Location = New-Object System.Drawing.Point(320, $y + 15)
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# Jalankan Form
$form.ShowDialog() | Out-Null
