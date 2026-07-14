$guiProcesses = Get-Process -Name powershell -ErrorAction SilentlyContinue | Where-Object {
    try {
        $cmdLine = (Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine
        $cmdLine -like "*VSSManager.ps1*"
    } catch {
        $false
    }
}

if ($guiProcesses) {
    "VSSManager is RUNNING! Processes:"
    foreach ($p in $guiProcesses) {
        "ID: $($p.Id)"
        $p.Kill()
    }
} else {
    "VSSManager is NOT running."
}
