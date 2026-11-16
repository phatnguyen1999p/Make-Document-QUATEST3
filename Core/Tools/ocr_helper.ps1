param(
    [string]$ListFile
)

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { 
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.IsGenericMethod 
} | Select-Object -First 1

function Await-WinRT($winRtTask, $resultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($resultType)
    $netTask = $asTask.Invoke($null, @($winRtTask))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}

[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($engine -eq $null) {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new("en-US"))
}

$paths = @()
if ($ListFile -and (Test-Path $ListFile)) {
    $paths = Get-Content $ListFile -Encoding UTF8
} else {
    $paths = $args
}

$results = [System.Collections.Generic.Dictionary[string, string]]::new()

foreach ($p in $paths) {
    $p = $p.Trim()
    if (-not $p) { continue }
    $fullPath = [System.IO.Path]::GetFullPath($p)
    if (-not (Test-Path $fullPath)) { 
        $results[$fullPath] = "FILE_NOT_FOUND"
        continue 
    }
    try {
        $fileOp = [Windows.Storage.StorageFile]::GetFileFromPathAsync($fullPath)
        $file = Await-WinRT $fileOp ([Windows.Storage.StorageFile])

        $streamOp = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
        $stream = Await-WinRT $streamOp ([Windows.Storage.Streams.IRandomAccessStream])

        $decoderOp = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
        $decoder = Await-WinRT $decoderOp ([Windows.Graphics.Imaging.BitmapDecoder])

        $bitmapOp = $decoder.GetSoftwareBitmapAsync()
        $bitmap = Await-WinRT $bitmapOp ([Windows.Graphics.Imaging.SoftwareBitmap])

        $ocrOp = $engine.RecognizeAsync($bitmap)
        $ocrRes = Await-WinRT $ocrOp ([Windows.Media.Ocr.OcrResult])

        $results[$fullPath] = $ocrRes.Text
    } catch {
        $results[$fullPath] = "ERROR: " + $_.ToString()
    }
}

$outFile = $ListFile -replace "\.txt$", "_out.json"
if ($ListFile -and $outFile -ne $ListFile) {
    $json = $results | ConvertTo-Json
    [System.IO.File]::WriteAllText($outFile, $json, [System.Text.Encoding]::UTF8)
    Write-Output "SAVED:$outFile"
} else {
    $json = $results | ConvertTo-Json
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Output $json
}
