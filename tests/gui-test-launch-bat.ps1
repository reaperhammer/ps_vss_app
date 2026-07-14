# gui-test-launch-bat.ps1
# Script to verify that Launch.bat successfully starts the GUI end-to-end.

$ErrorActionPreference = 'Stop'

Write-Host "Running Launch.bat..."
$process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c Launch.bat" -PassThru -WorkingDirectory "c:\project"

Write-Host "Waiting 6 seconds for Launch.bat to elevate, resolve directories, and launch the GUI..."
Start-Sleep -Seconds 6

# Take desktop screenshot
Write-Host "Taking desktop screenshot..."
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$bounds = $screen.Bounds
$bmp = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

$artifactPath = "C:\Users\will\.gemini\antigravity\brain\c49b84ec-69c1-4609-b013-adf62112dec1\gui_launch_bat_screenshot.png"
$bmp.Save($artifactPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bmp.Dispose()
Write-Host "Screenshot saved to $artifactPath"

# Check if VSSManager is running
Write-Host "Checking for VSSManager process..."
$guiProcesses = Get-Process -Name powershell -ErrorAction SilentlyContinue | Where-Object {
    try {
        $cmdLine = (Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine
        $cmdLine -like "*VSSManager.ps1*"
    } catch {
        $false
    }
}

if ($guiProcesses) {
    Write-Host "Success! Found $($guiProcesses.Count) running VSSManager GUI process(es)."
    foreach ($p in $guiProcesses) {
        Write-Host "Closing process $($p.Id)..."
        $p.CloseMainWindow() | Out-Null
        Start-Sleep -Seconds 1
        if (-not $p.HasExited) {
            $p.Kill()
        }
    }
} else {
    Write-Error "Verification FAILED: VSSManager process is NOT running after launching Launch.bat."
}
