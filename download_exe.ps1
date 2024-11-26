$exeUrl = "https://github.com/jojihatzz/combine_pdfs/raw/refs/heads/main/gui.exe"

$tempDir = New-Item -ItemType Directory -Path "$env:TEMP\MyTempFolder" -Force

$exePath = Join-Path -Path $tempDir.FullName -ChildPath "example_gui.exe"

Invoke-WebRequest -Uri $exeUrl -OutFile $exePath

$process = Start-Process -FilePath $exePath -PassThru
$process.WaitForExit()

Remove-Item -Path $exePath -Force

if (-Not (Get-ChildItem -Path $tempDir.FullName)) {
    Remove-Item -Path $tempDir.FullName -Force
}
