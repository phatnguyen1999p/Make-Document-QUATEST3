# chay_132115.ps1 - Chay ho so phong sach Q326D132115 (that, tren Input\132115).
# Thu tu: sao luu -> fill_excel_132115.py -> generate_all_reports.py (DOCX + PDF + UpChungThu) -> qa_check.
# - Sao luu Input\132115 vao %TEMP%\lamkq_backup_132115_<gio> truoc khi ghi de.
# - TUYET DOI KHONG dong bo o mang Z: (quy dinh cua du an): luon dat QUATEST3_NETWORK_ROOT tro vao thu muc khong ton tai.
# - Khong dong Word/Excel dang mo cua nguoi dung.
# Log: chay_132115_log.txt (cung thu muc). Gui file log cho Claude de doc.

$ErrorActionPreference = "Continue"
$src = $PSScriptRoot
$log = Join-Path $src "chay_132115_log.txt"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

if (Test-Path $log) { Remove-Item $log -Force }
function Log([string]$s) { Add-Content -Path $log -Value $s -Encoding UTF8; Write-Host $s }

Log "=== CHAY HO SO 132115 $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') ==="
Set-Location $src

# Chon Python: uu tien (1) bien LAMKQ_PYTHON, (2) Python cai theo user (%LOCALAPPDATA%\Programs\Python), (3) python/py trong PATH.
# Chi nhan ban nao import duoc cac thu vien du an (openpyxl, docx, fitz, lxml); bo qua ban thieu thu vien (vd Python he thong/admin).
$cands = @()
if ($env:LAMKQ_PYTHON) { $cands += $env:LAMKQ_PYTHON }
$userPyRoot = Join-Path $env:LOCALAPPDATA "Programs\Python"
if (Test-Path $userPyRoot) {
    $cands += @(Get-ChildItem -Path $userPyRoot -Directory -Filter "Python*" -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | ForEach-Object { Join-Path $_.FullName "python.exe" } | Where-Object { Test-Path $_ })
}
foreach ($c in @("python", "py")) { $g = Get-Command $c -ErrorAction SilentlyContinue; if ($g) { $cands += $g.Source } }
$py = $null
foreach ($c in ($cands | Select-Object -Unique)) {
    $ver = (& $c --version 2>&1 | Out-String).Trim()
    & $c -c "import openpyxl, docx, fitz, lxml" 2>$null
    if ($LASTEXITCODE -eq 0) { $py = $c; Log "Chon Python: $c ($ver)"; break }
    else { Log "Bo qua Python thieu thu vien: $c ($ver)" }
}
if (-not $py) { Log "[FATAL] Khong co Python nao du thu vien (openpyxl, python-docx, PyMuPDF, lxml). Dat bien LAMKQ_PYTHON tro toi python.exe dung."; exit 1 }
Log ("Python: " + (& $py --version 2>&1))

if (-not (Test-Path "Input\132115")) { Log "[FATAL] Khong thay Input\132115"; exit 1 }

$env:QUATEST3_NETWORK_ROOT = Join-Path $env:TEMP "NO_NETWORK_DRIVE_132115"
Log "Dong bo o mang Z: TAT (khong bao gio dong bo len Z:)."
# May in: KHONG dung may in tem HPL. pdf_exporter tu doi may in mac dinh khi can; o day ghi lai va tra lai neu Python bi ngat giua chung.
$printerBefore = (Get-CimInstance Win32_Printer | Where-Object { $_.Default }).Name
Log "May in mac dinh truoc khi chay: $printerBefore"
function Restore-Printer {
    $now = (Get-CimInstance Win32_Printer | Where-Object { $_.Default }).Name
    if ($printerBefore -and $now -ne $printerBefore) {
        Get-CimInstance Win32_Printer | Where-Object { $_.Name -eq $printerBefore } | Invoke-CimMethod -MethodName SetDefaultPrinter | Out-Null
        Log "Tra lai may in mac dinh: $printerBefore (dang la $now)"
    } else { Log "May in mac dinh sau khi chay: $now" }
}

# 1. Sao luu
$bk = Join-Path $env:TEMP ("lamkq_backup_132115_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
robocopy "Input\132115" $bk /E /NFL /NDL /NJH /NJS /NP | Out-Null
Log "Da sao luu Input\132115 -> $bk"

$wordBefore = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excelBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
Log "Word dang mo (PID): $($wordBefore -join ', ') | Excel dang mo (PID): $($excelBefore -join ', ')"
Log "Luu y: nen dong file .docx/.xlsx cua 132115 truoc khi chay de khong bi khoa file."

$results = @()
function Run-Step([string]$name, [string[]]$stepArgs) {
    Log ""
    Log "################ STEP: $name  ($($stepArgs -join ' ')) ################"
    $t0 = Get-Date
    # Ghi log tung dong (UTF-8, realtime). Khong dung Tee-Object: no day output vao gia tri tra ve cua ham
    # (Run-Step tra ve mang thay vi exit code -> bao loi nham) va ghi UTF-16 lan voi UTF-8.
    & $py @stepArgs 2>&1 | ForEach-Object { Log ([string]$_) }
    $code = $LASTEXITCODE
    $sec = [int]((Get-Date) - $t0).TotalSeconds
    Log "---- ket thuc STEP '$name': exit=$code, ${sec}s ----"
    $script:results += [pscustomobject]@{ Step = $name; Exit = $code; Seconds = $sec }
    return $code
}

try {
    $c1 = Run-Step "Do du lieu vao Excel tinh toan" @("fill_excel_132115.py")
    if ($c1 -ne 0) { Log "[DUNG] fill_excel that bai, khong chay tiep. Co the khoi phuc tu: $bk"; exit 1 }

    $c2 = Run-Step "Sinh BCKQ + TBKQ VN/EN + xuat PDF" @("generate_all_reports.py")
    if ($c2 -ne 0) { Log "[DUNG] generate_all_reports that bai. Co the khoi phuc tu: $bk"; exit 1 }

    Run-Step "QA final-check ThucHien" @("Core\Tools\qa_check.py", "final-check", "Input\132115\ThucHien", "--code", "Q326D132115") | Out-Null
    Run-Step "QA final-check TBKQ"     @("Core\Tools\qa_check.py", "final-check", "Input\132115\TBKQ", "--code", "Q326D132115") | Out-Null
    Run-Step "QA check-pdf"            @("Core\Tools\qa_check.py", "check-pdf", "Input\132115\pdf") | Out-Null

    $wordAfter = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $excelAfter = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    Log ""
    Log "Word truoc: [$($wordBefore -join ', ')] sau: [$($wordAfter -join ', ')]; Excel truoc: [$($excelBefore -join ', ')] sau: [$($excelAfter -join ', ')]"

    Log ""
    Log "################ FILE KET QUA ################"
    Get-ChildItem -Path "Input\132115" -Recurse -Include *.docx,*.xlsx,*.pdf -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch "Data Do|~\$" } |
        ForEach-Object { Log ("  " + $_.FullName.Substring($src.Length + 1) + "  (" + $_.Length + " bytes, " + $_.LastWriteTime.ToString("dd/MM HH:mm") + ")") }

    Log ""
    Log "################ TONG KET ################"
    foreach ($r in $results) { Log ("{0,-40} exit={1}  {2}s" -f $r.Step, $r.Exit, $r.Seconds) }
    Log "Xong. Ban sao luu: $bk"
    Log "Bao Claude: 'doc log 132115'."
} finally {
    Restore-Printer
    Set-Location $src
}
