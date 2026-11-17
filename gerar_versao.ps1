<#
.SYNOPSIS
    Gera uma nova versão do Monitor de Produtividade e deixa o instalador pronto para publicar.

.DESCRIPTION
    1. Atualiza a versão no app (APP_VERSAO) e no instalador (MeuAppVersao).
    2. Gera o executável (PyInstaller) e o instalador (Inno Setup).
    3. Copia o instalador para wwwroot\downloads do WebAppSystems e tira de lá as versões antigas
       (elas continuam guardadas na pasta Output).

    O WebAppSystems descobre a versão pelo nome do arquivo em downloads, então depois do script
    basta publicar o site: quem estiver com uma versão anterior recebe o aviso de atualização.

.EXAMPLE
    .\gerar_versao.ps1 1.0.1
#>
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Versao,

    # Pasta wwwroot\downloads do WebAppSystems
    [string]$PastaDownloads = "C:\temp\Projeto\WebAppSystems\WebAppSystems\wwwroot\downloads",

    # Mantém as versões antigas em downloads (por padrão são removidas de lá)
    [switch]$ManterAntigos
)

$ErrorActionPreference = "Stop"
$pasta = $PSScriptRoot
$arquivoApp = Join-Path $pasta "app_timesheet_integrado.py"
$arquivoIss = Join-Path $pasta "instalador_timesheet.iss"
$iscc = "C:\Program Files\Inno Setup 7\ISCC.exe"
$utf8SemBom = New-Object System.Text.UTF8Encoding($false)

function Passo($texto) { Write-Host "`n==> $texto" -ForegroundColor Cyan }
function Falha($texto) { Write-Host "`nERRO: $texto" -ForegroundColor Red; exit 1 }

# ---------- Validações ----------
if ($Versao -notmatch '^\d+\.\d+\.\d+$') { Falha "Versão inválida '$Versao'. Use o formato 1.0.1" }
foreach ($f in $arquivoApp, $arquivoIss, $iscc) { if (-not (Test-Path $f)) { Falha "Não encontrado: $f" } }
if (-not (Test-Path $PastaDownloads)) { Falha "Pasta de downloads não encontrada: $PastaDownloads" }

$textoApp = [IO.File]::ReadAllText($arquivoApp, $utf8SemBom)
$atual = [regex]::Match($textoApp, '(?m)^APP_VERSAO\s*=\s*"([^"]+)"').Groups[1].Value
if (-not $atual) { Falha "APP_VERSAO não encontrado em $arquivoApp" }
if ([version]$Versao -le [version]$atual) {
    Falha "A nova versão ($Versao) precisa ser maior que a atual ($atual). Quem já tem a $atual não seria avisado."
}
Write-Host "Versão atual: $atual  ->  nova: $Versao"

# ---------- 1) Versão no app e no instalador ----------
Passo "Atualizando a versão no app e no instalador"
$textoApp = [regex]::Replace($textoApp, '(?m)^APP_VERSAO\s*=\s*"[^"]+"', "APP_VERSAO = `"$Versao`"")
[IO.File]::WriteAllText($arquivoApp, $textoApp, $utf8SemBom)

$textoIss = [IO.File]::ReadAllText($arquivoIss, $utf8SemBom)
if ($textoIss -notmatch '#define MeuAppVersao "[^"]+"') { Falha "MeuAppVersao não encontrado em $arquivoIss" }
$textoIss = [regex]::Replace($textoIss, '#define MeuAppVersao "[^"]+"', "#define MeuAppVersao `"$Versao`"")
[IO.File]::WriteAllText($arquivoIss, $textoIss, $utf8SemBom)

& python -m py_compile $arquivoApp
if ($LASTEXITCODE -ne 0) { Falha "O app não compila (py_compile)." }

# ---------- 2) Executável e instalador ----------
Passo "Gerando o executável (PyInstaller). Leva cerca de 1 a 2 minutos..."
Push-Location $pasta
try {
    $ErrorActionPreference = "Continue"     # o PyInstaller escreve o progresso no stderr
    & python -m PyInstaller app_timesheet_integrado.spec --noconfirm --clean --log-level WARN
    $codigo = $LASTEXITCODE
    $ErrorActionPreference = "Stop"
    if ($codigo -ne 0) { Falha "PyInstaller terminou com erro ($codigo)." }

    Passo "Gerando o instalador (Inno Setup)"
    & $iscc /Q instalador_timesheet.iss
    if ($LASTEXITCODE -ne 0) { Falha "Inno Setup terminou com erro ($LASTEXITCODE)." }
}
finally { Pop-Location }

$nomeInstalador = "MonitorTimeSheet-Setup-$Versao.exe"
$instalador = Join-Path $pasta "Output\$nomeInstalador"
if (-not (Test-Path $instalador)) { Falha "Instalador não foi gerado: $instalador" }

# ---------- 3) Copia para o WebAppSystems ----------
Passo "Copiando para $PastaDownloads"
Copy-Item $instalador $PastaDownloads -Force
if (-not $ManterAntigos) {
    Get-ChildItem $PastaDownloads -Filter "MonitorTimeSheet-Setup-*.exe" |
        Where-Object { $_.Name -ne $nomeInstalador } |
        ForEach-Object { Remove-Item $_.FullName -Force; Write-Host "  removido de downloads: $($_.Name) (continua em Output)" }
}

$hashOrigem = (Get-FileHash $instalador -Algorithm SHA256).Hash
$hashDestino = (Get-FileHash (Join-Path $PastaDownloads $nomeInstalador) -Algorithm SHA256).Hash
if ($hashOrigem -ne $hashDestino) { Falha "A cópia em downloads ficou diferente do instalador gerado." }

$tamanho = [math]::Round((Get-Item $instalador).Length / 1MB, 1)
Write-Host "`nPRONTO: versão $Versao" -ForegroundColor Green
Write-Host "  Instalador: $instalador ($tamanho MB)"
Write-Host "  SHA-256:    $hashOrigem"
Write-Host "  Copiado em: $PastaDownloads"
Write-Host "`nPróximo passo: publicar o WebAppSystems. Quem tiver uma versão anterior recebe o aviso ao abrir o Monitor." -ForegroundColor Yellow
