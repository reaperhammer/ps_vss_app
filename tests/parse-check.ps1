$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '../VSSManager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) {
    $errors | Select-Object -First 10 | ForEach-Object { "Line $($_.Extent.StartLineNumber): $($_.Message)" }
    exit 1
} else {
    'No errors'
    exit 0
}
