
try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
    [void][Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]
    [void][Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime]
    [void][Windows.Storage.Streams.IRandomAccessStream, Windows.Foundation, ContentType = WindowsRuntime]
    [void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType = WindowsRuntime]
    [void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]

    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { 
        $_.Name -eq 'AsTask' -and 
        $_.GetParameters().Count -eq 1 -and 
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' 
    })[0]

    function Await-WinRT($asyncOperation, $resultType) {
        $asTask = $asTaskGeneric.MakeGenericMethod($resultType)
        $netTask = $asTask.Invoke($null, @($asyncOperation))
        $netTask.Wait(-1) | Out-Null
        return $netTask.Result
    }

    $lang = New-Object Windows.Globalization.Language("zh-Hant-TW")
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
    
    $absPath = [System.IO.Path]::GetFullPath($args[0])
    $fileOperation = [Windows.Storage.StorageFile]::GetFileFromPathAsync($absPath)
    $file = Await-WinRT $fileOperation ([Windows.Storage.StorageFile])

    $streamOperation = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
    $stream = Await-WinRT $streamOperation ([Windows.Storage.Streams.IRandomAccessStream])

    $decoderOperation = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
    $decoder = Await-WinRT $decoderOperation ([Windows.Graphics.Imaging.BitmapDecoder])

    $bitmapOperation = $decoder.GetSoftwareBitmapAsync()
    $softwareBitmap = Await-WinRT $bitmapOperation ([Windows.Graphics.Imaging.SoftwareBitmap])

    $ocrOperation = $engine.RecognizeAsync($softwareBitmap)
    $result = Await-WinRT $ocrOperation ([Windows.Media.Ocr.OcrResult])

    Write-Output $result.Text
}
catch {
    Write-Error $_.Exception.Message
}
