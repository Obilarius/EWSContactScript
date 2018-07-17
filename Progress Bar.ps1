### Progress Bar

for ($a = 1; $a -le 100; $a++) {
    Write-Progress -Activity "Kontakte werden übertragen..." -PercentComplete $a -CurrentOperation "$a% complete" -Status "Please wait"
    Start-Sleep 1
}