$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseDir = Resolve-Path (Join-Path $scriptDir "..")
Set-Location $baseDir

# Forca UTF-8 na sessao (evita acentos quebrados)
chcp 65001 | Out-Null
$OutputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

# SharePoint site URL (ajuste se necessario)
$env:SP_SITE_URL = "https://casacredsacombr.sharepoint.com/sites/Processos"
$env:SP_LIBRARY_NAME = "Comercial"

# Pasta de destino dentro do site
$env:SP_FOLDER_PATH = "1. Analise de Credito"

$required = @("SP_TENANT_ID", "SP_CLIENT_ID", "SP_CLIENT_SECRET", "SP_SITE_URL")
$missing = @()
foreach ($name in $required) {
    if (-not (Get-Item -Path "Env:$name" -ErrorAction SilentlyContinue)) { $missing += $name }
}
$logFile = Join-Path $baseDir ("05_LOGS\\run_click_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

if ($missing.Count -gt 0) {
    "Faltam variaveis de ambiente: $($missing -join ', ')" | Tee-Object -FilePath $logFile -Append
    "Defina essas variaveis e rode novamente." | Tee-Object -FilePath $logFile -Append
    Exit 1
}

"Rodando pipeline em: $baseDir" | Tee-Object -FilePath $logFile -Append
"Log: $logFile" | Tee-Object -FilePath $logFile -Append

python .\02_SCRIPTS\00_RUN_ALL.py | Tee-Object -FilePath $logFile -Append

if ($LASTEXITCODE -eq 0) {
    "Concluido com sucesso." | Tee-Object -FilePath $logFile -Append
} else {
    "Finalizado com erros. Verifique o log." | Tee-Object -FilePath $logFile -Append
}
