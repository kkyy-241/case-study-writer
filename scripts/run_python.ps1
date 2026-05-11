param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PythonArgs
)

$ErrorActionPreference = "Stop"

$candidate = if ($env:CASE_WRITER_PYTHON) {
    $env:CASE_WRITER_PYTHON
} else {
    "D:\anaconda3\envs\python313\python.exe"
}

if (-not (Test-Path -LiteralPath $candidate)) {
    throw "Python interpreter not found: $candidate. Set CASE_WRITER_PYTHON to the Python executable for this repository."
}

& $candidate @PythonArgs
exit $LASTEXITCODE
