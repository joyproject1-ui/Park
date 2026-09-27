<#
FDA Warning Letter 수집·발송 — Windows 작업 스케줄러용 실행 스크립트.

    .\run_fda_letters.ps1 -DryRun        메일 없이 결과만 확인
    .\run_fda_letters.ps1                실제 발송
    .\run_fda_letters.ps1 -SterileOnly   무균·주사제·점안제 관련 건만

설정은 같은 폴더의 fda-mail.config.ps1 에서 읽습니다 (example 파일 참조).
실행 기록은 저장소 루트의 logs\ 에 날짜별로 남습니다.
#>
param(
    [switch]$DryRun,
    [switch]$SterileOnly,
    [int]$Since = 7
)

$ErrorActionPreference = 'Stop'

$root   = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$config = Join-Path $PSScriptRoot 'fda-mail.config.ps1'

if (-not (Test-Path $config)) {
    Write-Host "설정 파일이 없습니다: $config" -ForegroundColor Red
    Write-Host "fda-mail.config.example.ps1 을 복사해 fda-mail.config.ps1 로 저장하고 값을 채우세요."
    exit 2
}
. $config

# 스케줄러로 돌 때 콘솔이 없어도 한글 출력이 깨지지 않도록 합니다.
$env:PYTHONUTF8       = '1'
$env:PYTHONIOENCODING = 'utf-8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$logDir = Join-Path $root 'logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir ('fda-letters-{0:yyyyMMdd-HHmm}.log' -f (Get-Date))

$pyArgs = @('-m', 'gmpai', 'letters', '--since', "$Since")
if ($SterileOnly) { $pyArgs += '--sterile-only' }
if (-not $DryRun) { $pyArgs += '--mail' }

function Log([string]$line) {
    Write-Host $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

Set-Location $root
Log "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] python $($pyArgs -join ' ')"

# Windows PowerShell 5.1 은 'Stop' 상태에서 stderr 를 오류로 승격시키므로 실행 중에는 풀어 둡니다.
$ErrorActionPreference = 'Continue'
& python @pyArgs 2>&1 | ForEach-Object { Log "$_" }
$code = $LASTEXITCODE
Log "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 종료 코드 $code"

# 30일 지난 기록은 정리
Get-ChildItem $logDir -Filter 'fda-letters-*.log' |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
    Remove-Item -ErrorAction SilentlyContinue

exit $code
