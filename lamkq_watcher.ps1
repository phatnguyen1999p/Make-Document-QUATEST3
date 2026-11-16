# lamkq_watcher.ps1 - Tien trinh nen nho: cho file lenh lamkq_trigger.txt roi tu chay chay_<ma>.ps1.
# Muc dich: Claude kich hoat chay ho so ma KHONG phai dung chuot/ban phim vao File Explorer cua nguoi dung.
# An toan:
#  - Chi chay chay_<so>.ps1 NAM TRONG CUNG THU MUC voi file nay (ma ho so chi gom chu so).
#  - Khong dong bo o mang Z: (chay_<ma>.ps1 da chan Z:). Khong ghi ra ngoai thu muc cho phep.
#  - Chi mot ban chay (mutex). Dong cua so/ tien trinh powershell nay la dung.
# Cach dung: bam dup start_lamkq_watcher.bat 1 lan sau khi khoi dong may. Dung: stop_lamkq_watcher.bat.
# Luu y: duong dan co dau cach (vd "Lam KQ") nen phai boc $script trong dau ngoac kep khi truyen cho Start-Process.
# Trang thai: lamkq_status.txt (RUNNING/DONE/ERROR), nhip tim: lamkq_watcher_alive.txt, log chay: chay_<ma>_log.txt

$ErrorActionPreference = "Continue"
$dir = $PSScriptRoot
Set-Location $dir
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, ("Global\LamKQ_Watcher_" + ($dir -replace "[^A-Za-z0-9]", "_")), [ref]$createdNew)
if (-not $createdNew) { exit 0 }

$trigger = Join-Path $dir "lamkq_trigger.txt"
$stop    = Join-Path $dir "lamkq_stop.txt"
$status  = Join-Path $dir "lamkq_status.txt"
$alive   = Join-Path $dir "lamkq_watcher_alive.txt"
if (Test-Path $stop) { Remove-Item $stop -Force -ErrorAction SilentlyContinue }

function Set-Status([string]$s) { Set-Content -Path $status -Value $s -Encoding UTF8 }
$lastBeat = [DateTime]::MinValue
Set-Status ("IDLE watcher started " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss"))

while ($true) {
    if (Test-Path $stop) { Set-Status ("STOPPED " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss")); Remove-Item $stop -Force -ErrorAction SilentlyContinue; break }
    if (((Get-Date) - $lastBeat).TotalSeconds -ge 5) {
        Set-Content -Path $alive -Value (Get-Date -Format "dd/MM/yyyy HH:mm:ss") -Encoding ASCII
        $lastBeat = Get-Date
    }
    if (Test-Path $trigger) {
        Start-Sleep -Milliseconds 800   # cho ghi xong file
        $raw = ""
        try { $raw = (Get-Content -Path $trigger -Raw -ErrorAction Stop) } catch { Start-Sleep -Seconds 1; continue }
        Remove-Item $trigger -Force -ErrorAction SilentlyContinue
        $id = ($raw -replace "[^0-9]", "")
        $script = Join-Path $dir ("chay_" + $id + ".ps1")
        if ($id -eq "" -or $id.Length -gt 10 -or -not (Test-Path $script)) {
            Set-Status ("ERROR trigger khong hop le: '" + $raw.Trim() + "' (can chay_<so>.ps1 trong thu muc) " + (Get-Date -Format "HH:mm:ss"))
            continue
        }
        $t0 = Get-Date
        Set-Status ("RUNNING " + $id + " bat dau " + $t0.ToString("dd/MM/yyyy HH:mm:ss"))
        $p = Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ("`"" + $script + "`"")) `
             -WorkingDirectory $dir -WindowStyle Hidden -PassThru -Wait
        $sec = [int]((Get-Date) - $t0).TotalSeconds
        $state = if ($p.ExitCode -eq 0) { "DONE" } else { "ERROR" }
        Set-Status ("$state " + $id + " exit=" + $p.ExitCode + " " + $sec + "s ket thuc " + (Get-Date -Format "dd/MM/yyyy HH:mm:ss") + " (log: chay_" + $id + "_log.txt)")
    }
    Start-Sleep -Seconds 2
}
$mutex.ReleaseMutex()
