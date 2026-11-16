# OCR anh bang Windows.Media.Ocr, tra ve tung tu kem toa do (de dung lai bo cuc bang).
# Dung: powershell -File ocr_boxes.ps1 -ListFile <list.txt>  -> ghi <list>_out.json
param([string]$ListFile)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.IsGenericMethod } | Select-Object -First 1
function Await-WinRT($t, $type) { $n = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($t)); $n.Wait(-1) | Out-Null; $n.Result }
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new("en-US"))
if ($engine -eq $null) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
$out = @{}
foreach ($p in (Get-Content $ListFile -Encoding UTF8)) {
    $p = $p.Trim(); if (-not $p) { continue }
    try {
        $f = Await-WinRT ([Windows.Storage.StorageFile]::GetFileFromPathAsync($p)) ([Windows.Storage.StorageFile])
        $s = Await-WinRT ($f.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        $d = Await-WinRT ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($s)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $b = Await-WinRT ($d.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $r = Await-WinRT ($engine.RecognizeAsync($b)) ([Windows.Media.Ocr.OcrResult])
        $words = @()
        foreach ($ln in $r.Lines) { foreach ($w in $ln.Words) {
            $words += ,@($w.Text, [math]::Round($w.BoundingRect.X,1), [math]::Round($w.BoundingRect.Y,1), [math]::Round($w.BoundingRect.Width,1), [math]::Round($w.BoundingRect.Height,1))
        } }
        $out[$p] = @{ w = $b.PixelWidth; h = $b.PixelHeight; words = $words }
        $s.Dispose()
    } catch { $out[$p] = @{ error = $_.ToString() } }
}
$json = $out | ConvertTo-Json -Depth 6 -Compress
[System.IO.File]::WriteAllText(($ListFile -replace '\.txt$', '_out.json'), $json, [System.Text.UTF8Encoding]::new($false))
