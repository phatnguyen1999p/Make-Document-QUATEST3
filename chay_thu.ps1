# chay_thu.ps1 - Chay thu toan bo ho so mau (co xuat PDF qua Word/Excel COM) tren BAN SAO cua du an.
# - Khong dong vao Input\, Output\, o mang Z:. Moi thu chay trong %TEMP%\lamkq_chay_thu.
# - Ket qua ghi vao chay_thu_log.txt (cung thu muc voi script nay). Gui file log cho Claude de doc.
# Cach chay: bam dup chay_thu.bat (hoac: powershell -ExecutionPolicy Bypass -File chay_thu.ps1)

$ErrorActionPreference = "Continue"
$src  = $PSScriptRoot
$log  = Join-Path $src "chay_thu_log.txt"
$work = Join-Path $env:TEMP "lamkq_chay_thu"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

if (Test-Path $log) { Remove-Item $log -Force }
function Log([string]$s) { Add-Content -Path $log -Value $s -Encoding UTF8; Write-Host $s }

Log "=== CHAY THU $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') ==="

# Python
$py = $null
foreach ($c in @("python", "py")) { if (Get-Command $c -ErrorAction SilentlyContinue) { $py = $c; break } }
if (-not $py) { Log "[FATAL] Khong tim thay python/py trong PATH."; exit 1 }
Log ("Python: " + (& $py --version 2>&1))

# Ban sao lam viec
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory -Path $work | Out-Null
robocopy $src $work /E /XD __pycache__ dist quatest3_core.egg-info Output scratch .git /XF layout_crop_*.png receipt_*.png scan_p*.png chay_thu_log.txt /MAX:50000000 /NFL /NDL /NJH /NJS /NP | Out-Null
if (-not (Test-Path (Join-Path $work "Input"))) { New-Item -ItemType Directory -Path (Join-Path $work "Input") | Out-Null }
Log "Ban sao lam viec: $work"

# Chan dong bo o mang (tranh ghi de ho so that tren Z:)
$env:QUATEST3_NETWORK_ROOT = Join-Path $work "NO_NETWORK_DRIVE"

# Kiem tra E2: Word dang mo truoc khi chay (mo san 1 file Word chua luu de thu)
$wordBefore = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excelBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
Log "Word dang mo truoc khi chay (PID): $($wordBefore -join ', ')"
Log "Excel dang mo truoc khi chay (PID): $($excelBefore -join ', ')"

Set-Location $work
$results = @()

function Run-Step([string]$name, [string[]]$stepArgs) {
    Log ""
    Log "################ STEP: $name  ($($stepArgs -join ' ')) ################"
    $t0 = Get-Date
    $out = & $py "Core\run.py" @stepArgs 2>&1
    $code = $LASTEXITCODE
    foreach ($l in $out) { Add-Content -Path $log -Value ([string]$l) -Encoding UTF8 }
    $sec = [int]((Get-Date) - $t0).TotalSeconds
    Log "---- ket thuc STEP '$name': exit=$code, ${sec}s ----"
    $script:results += [pscustomobject]@{ Step = $name; Exit = $code; Seconds = $sec }
}

Run-Step "BSC mau"                 @("--sample-bsc")
Run-Step "Tu hut mau"              @("--sample-fumehood")
Run-Step "Tu hut - ca 2 form"      @("--sample-fumehood", "--both-forms")
Run-Step "LAF mau"                 @("--sample-laf")
Run-Step "Khi nen mau"             @("--sample-khinen")
Run-Step "Khi nen 133266 (KN1+KN2)" @("--sample-khinen-133266")

# QA tren toan bo ket qua
Log ""
Log "################ QA final-check (toan bo Input\ cua ban sao) ################"
$qa = & $py "Core\Tools\qa_check.py" final-check "Input" 2>&1
foreach ($l in $qa) { Add-Content -Path $log -Value ([string]$l) -Encoding UTF8 }

# Mo thu qua Word COM, dem trang
Log ""
Log "################ QA test-open (Word COM) ################"
$to = & $py "Core\Tools\qa_check.py" test-open "Input" 2>&1
foreach ($l in $to) { Add-Content -Path $log -Value ([string]$l) -Encoding UTF8 }

# Kiem tra E2: Word/Excel cua nguoi dung co con song khong
$wordAfter = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excelAfter = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$wordLost = @($wordBefore | Where-Object { $wordAfter -notcontains $_ })
$excelLost = @($excelBefore | Where-Object { $excelAfter -notcontains $_ })
Log ""
Log "################ KIEM TRA E2 (khong dong Word/Excel cua nguoi dung) ################"
Log "Word truoc: [$($wordBefore -join ', ')]  sau: [$($wordAfter -join ', ')]  bi dong: [$($wordLost -join ', ')]"
Log "Excel truoc: [$($excelBefore -join ', ')]  sau: [$($excelAfter -join ', ')]  bi dong: [$($excelLost -join ', ')]"
if ($wordBefore.Count -eq 0 -and $excelBefore.Count -eq 0) { Log "(Luu y: luc bat dau khong co Word/Excel nao mo, nen chua chung minh duoc E2. Mo san 1 file Word roi chay lai.)" }

# Thong ke file ket qua
Log ""
Log "################ FILE KET QUA ################"
$pdfs = @(Get-ChildItem -Path (Join-Path $work "Input") -Recurse -Filter *.pdf -ErrorAction SilentlyContinue)
$docx = @(Get-ChildItem -Path (Join-Path $work "Input") -Recurse -Filter *.docx -ErrorAction SilentlyContinue)
$xlsx = @(Get-ChildItem -Path (Join-Path $work "Input") -Recurse -Filter *.xlsx -ErrorAction SilentlyContinue)
Log "DOCX: $($docx.Count)   XLSX: $($xlsx.Count)   PDF: $($pdfs.Count)"
foreach ($p in $pdfs) { Log ("  PDF " + $p.FullName.Substring($work.Length) + "  (" + $p.Length + " bytes)") }

Log ""
Log "################ TONG KET ################"
foreach ($r in $results) { Log ("{0,-28} exit={1}  {2}s" -f $r.Step, $r.Exit, $r.Seconds) }
Log "Xong. Gui file chay_thu_log.txt cho Claude. Ban sao ket qua nam o: $work"
Set-Location $src
