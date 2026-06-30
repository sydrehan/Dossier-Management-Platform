# Kill any running Word instances first
Stop-Process -Name "WINWORD" -Force -ErrorAction SilentlyContinue

# Clear Word's Resiliency registry keys (which store crash state and trigger Safe Mode prompts)
$resilityPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency",
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DocumentRecovery",
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\StartupItems"
)

foreach ($path in $resilityPaths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Clear AutoRecovery files
$autoRecoveryPath = "$env:APPDATA\Microsoft\Word"
if (Test-Path $autoRecoveryPath) {
    Remove-Item -Path "$autoRecoveryPath\*.asd" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$autoRecoveryPath\*.*" -Include "*~$*" -Force -ErrorAction SilentlyContinue
}

Write-Output "Word resiliency registry keys and autorecovery files cleared successfully."
