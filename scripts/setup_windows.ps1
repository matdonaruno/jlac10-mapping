# ============================================================
# JLAC10 MAPPER - Windows 環境セットアップスクリプト
# ============================================================
# PowerShell で実行: .\scripts\setup_windows.ps1
# 管理者権限不要（winget / py コマンドを使用）
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " JLAC10 MAPPER - Environment Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ------------------------------------------------------------
# 1. Python チェック & インストール
# ------------------------------------------------------------
Write-Host "[1/4] Python ..." -ForegroundColor Yellow

$pyVersion = $null
try {
    $pyVersion = & py --version 2>&1
} catch {}

if ($pyVersion -match "Python \d+\.\d+") {
    Write-Host "  OK: $pyVersion" -ForegroundColor Green
} else {
    Write-Host "  Python が見つかりません。インストールします..." -ForegroundColor Red
    winget install Python.Python.3.12
    Write-Host "  インストール完了。PowerShell を再起動してから再度実行してください。" -ForegroundColor Yellow
    exit 1
}

# ------------------------------------------------------------
# 2. uv チェック & インストール
# ------------------------------------------------------------
Write-Host "[2/4] uv (Python package manager) ..." -ForegroundColor Yellow

$uvVersion = $null
try {
    $uvVersion = & uv --version 2>&1
} catch {}

if ($uvVersion -match "uv \d+") {
    Write-Host "  OK: $uvVersion" -ForegroundColor Green
} else {
    Write-Host "  uv をインストールします..." -ForegroundColor Cyan
    py -m pip install uv --quiet
    $uvVersion = & py -m uv --version 2>&1
    Write-Host "  OK: $uvVersion" -ForegroundColor Green
}

# ------------------------------------------------------------
# 3. Git チェック & インストール
# ------------------------------------------------------------
Write-Host "[3/4] Git ..." -ForegroundColor Yellow

$gitVersion = $null
try {
    $gitVersion = & git --version 2>&1
} catch {}

if ($gitVersion -match "git version") {
    Write-Host "  OK: $gitVersion" -ForegroundColor Green
} else {
    Write-Host "  Git をインストールします..." -ForegroundColor Cyan
    winget install Git.Git
    Write-Host "  インストール完了。PowerShell を再起動してから再度実行してください。" -ForegroundColor Yellow
    exit 1
}

# ------------------------------------------------------------
# 4. プロジェクト依存関係インストール
# ------------------------------------------------------------
Write-Host "[4/4] プロジェクト依存関係 ..." -ForegroundColor Yellow

if (Test-Path "pyproject.toml") {
    py -m uv sync 2>&1 | Out-Null
    Write-Host "  OK: 依存関係インストール完了" -ForegroundColor Green
} else {
    Write-Host "  pyproject.toml が見つかりません。プロジェクトルートで実行してください。" -ForegroundColor Red
    exit 1
}

# ------------------------------------------------------------
# GHE SSL 設定（オンプレ GitHub 用）
# ------------------------------------------------------------
if (Test-Path ".git") {
    git config http.sslVerify false 2>&1 | Out-Null
    Write-Host ""
    Write-Host "  GHE SSL検証を無効化しました（オンプレ用）" -ForegroundColor DarkGray
}

# ------------------------------------------------------------
# 完了
# ------------------------------------------------------------
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " Setup Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "使用方法:" -ForegroundColor Cyan
Write-Host ""
Write-Host "  py -m uv run srl-scraper --help       全コマンド一覧" -ForegroundColor White
Write-Host "  py -m uv run srl-scraper search TP     検索" -ForegroundColor White
Write-Host "  py -m uv run srl-scraper -v map ...    一括マッピング(デバッグ)" -ForegroundColor White
Write-Host ""
