Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- 1. PROSES PENGUMPULAN DATA HARDWARE ---
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


# --- 2. PEMBUATAN INTERFACE GUI LAYOUT MANUAL (ANTI-BUG) ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "System Info - Stock Opname Tool v2.1"
$form.Size = New-Object System.Drawing.Size(600, 560)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Font Global
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$fontText = New-Object System.Drawing.Font("Segoe UI", 10)

# Judul Utama
$title = New-Object System.Windows.Forms.Label
$title.Text = "LAPTOP SPECIFICATION FOR STOCK OPNAME"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::FromArgb(0, 173, 181)
$title.Size = New-Object System.Drawing.Size(540, 30)
$title.Location = New-Object System.Drawing.Point(20, 20)
$title.TextAlign = "ContentAlignment.MiddleCenter"
$form.Controls.Add($title)

# Fungsi Helper untuk membuat baris label dan textbox secara pas
function Add-Row ($labelText, $valueText, $yPos) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelText
    $lbl.Font = $fontLabel
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.Location = New-Object System.Drawing.Point(30, $yPos)
    $lbl.Size = New-Object System.Drawing.Size(140, 25)
    $lbl.TextAlign = "ContentAlignment.MiddleLeft"
    $form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text = $valueText
    $txt.Font = $fontText
    $txt.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $txt.ForeColor = [System.Drawing.Color]::White
    $txt.BorderStyle = "FixedSingle"
    $txt.Location = New-Object System.Drawing.Point(180, $yPos)
    $txt.Size = New-Object System.Drawing.Size(370, 25)
    $txt.ReadOnly = $true
    $form.Controls.Add($txt)
}

# Gambar komponen satu per satu dengan jarak Y yang aman (jarak 45px)
Add-Row "Computer Name" $CompName 70
Add-Row "Merk / Tipe" $Brand 115
Add-Row "OS Version" $OS 160
Add-Row "Processor" $CPU 205
Add-Row "RAM Capacity" $RAM 250
Add-Row "Storage Info" $Storage 295
Add-Row "Display / Grafis" $Layar 340
Add-Row "Serial Number" $SN 385

# --- TOMBOL-TOMBOL DI BAGIAN BAWAH ---
$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "📋 Copy All Data"
$btnCopy.Font = $fontLabel
$btnCopy.BackColor = [System.Drawing.Color]::FromArgb(0, 173, 181)
$btnCopy.ForeColor = [System.Drawing.Color]::White
$btnCopy.FlatStyle = "Flat"
$btnCopy.Size = New-Object System.Drawing.Size(160, 40)
$btnCopy.Location = New-Object System.Drawing.Point(120, 450)
$btnCopy.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCopy.FlatAppearance.BorderSize = 0
$btnCopy.Add_Click({
    $textToCopy = "Computer Name: $CompName`r`nMerk / Tipe: $Brand`r`nOS Version: $OS`r`nProcessor: $CPU`r`nRAM Capacity: $RAM`r`nStorage Info: $Storage`r`nDisplay / Grafis: $Layar`r`nSerial Number: $SN"
    [System.Windows.Forms.Clipboard]::SetText($textToCopy)
    [System.Windows.Forms.MessageBox]::Show("Semua data spesifikasi berhasil disalin ke clipboard!", "Sukses", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})
$form.Controls.Add($btnCopy)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "❌ Close"
$btnClose.Font = $fontLabel
$btnClose.BackColor = [System.Drawing.Color]::FromArgb(217, 83, 79)
$btnClose.ForeColor = [System.Drawing.Color]::White
$btnClose.FlatStyle = "Flat"
$btnClose.Size = New-Object System.Drawing.Size(160, 40)
$btnClose.Location = New-Object System.Drawing.Point(300, 450)
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# Jalankan Form Aplikasi
$form.ShowDialog() | Out-Null
