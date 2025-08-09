# BOOTH CSV 自動スクリーンショット撮影スクリプト（PowerShell 7対応版）
# 使用方法: scripts/screenshot フォルダから実行
# 例: pwsh .\capture-all.ps1

param(
    [string]$OutputDir,
    [string]$HtmlFile,
    [string]$ChromeBinaryPath
)

if (-not $OutputDir -or [string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = "../../docs/images" }
if (-not $HtmlFile -or [string]::IsNullOrWhiteSpace($HtmlFile)) { $HtmlFile = "../../boothcsv.html" }
if ($ChromeBinaryPath) { $ChromeBinaryPath = (Resolve-Path -LiteralPath $ChromeBinaryPath -ErrorAction SilentlyContinue) }

# PowerShell バージョン確認
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "⚠️  この スクリプトはPowerShell 7以降で実行してください" -ForegroundColor Yellow
    Write-Host "現在のバージョン: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "PowerShell 7で実行: pwsh .\capture-all.ps1" -ForegroundColor Cyan
    exit 1
}

# Seleniumモジュール確認
if (!(Get-Module -ListAvailable -Name Selenium)) {
    Write-Host "❌ Seleniumモジュールがインストールされていません" -ForegroundColor Red
    Write-Host "管理者権限でPowerShellを実行し、以下を実行してください:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name Selenium -Force" -ForegroundColor White
    exit 1
}

Import-Module Selenium


# 事前準備
$OutputDir = Resolve-Path $OutputDir -ErrorAction SilentlyContinue
if (!$OutputDir) {
    # 相対パスが解決できない場合は絶対パスで作成
    $OutputDir = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "docs\images"
}

if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$HtmlFile = Resolve-Path $HtmlFile -ErrorAction SilentlyContinue
if (!$HtmlFile) {
    # 相対パスが解決できない場合は絶対パスで構築
    $HtmlFile = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "boothcsv.html"
}

if (!(Test-Path $HtmlFile)) {
    Write-Host "❌ HTMLファイルが見つかりません: $HtmlFile" -ForegroundColor Red
    exit 1
}

# ChromeDriverの自動取得・更新
$chromeDriverDir = Get-Location
$chromeDriverExe = Join-Path $chromeDriverDir "chromedriver.exe"

function Get-ChromeVersion {
    param([string]$ExplicitPath)
    if ($ExplicitPath -and (Test-Path $ExplicitPath)) {
        try { return (Get-Item $ExplicitPath).VersionInfo.ProductVersion } catch { }
    }
    $chromePaths = @(
        "$env:ProgramFiles\\Google\\Chrome\\Application\\chrome.exe",
        "$env:ProgramFiles(x86)\\Google\\Chrome\\Application\\chrome.exe",
        "$env:LocalAppData\\Google\\Chrome\\Application\\chrome.exe"
    )
    foreach ($path in $chromePaths) {
        if (Test-Path $path) {
            $ver = (Get-Item $path).VersionInfo.ProductVersion
            if ($ver) { return $ver }
        }
    }
    return $null
}

function Get-ChromeDriverVersion {
    param($exePath)
    if (!(Test-Path $exePath)) { return $null }
    try {
        $output = & $exePath --version 2>$null
        if ($output -match 'ChromeDriver ([\d.]+)') {
            return $matches[1]
        }
    } catch {}
    return $null
}

function Download-ChromeDriver {
    param($version, $destPath)
    $major = $version.Split('.')[0]
    $url = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/$version/win64/chromedriver-win64.zip"
    $guid = [guid]::NewGuid().ToString()
    $tmpZip = Join-Path $env:TEMP ("chromedriver_" + $guid + ".zip")
    $tmpExtract = Join-Path $env:TEMP ("chromedriver_extract_" + $guid)
    Write-Host "🌐 ChromeDriver $version をダウンロード中..." -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpZip -UseBasicParsing
        if (Test-Path $tmpExtract) { Remove-Item $tmpExtract -Recurse -Force }
        Expand-Archive -Path $tmpZip -DestinationPath $tmpExtract -Force
        $driverPath = Get-ChildItem -Path $tmpExtract -Recurse -Filter "chromedriver.exe" | Select-Object -First 1
        if ($driverPath) {
            Copy-Item $driverPath.FullName -Destination $destPath -Force
            Write-Host "✅ ChromeDriver $version を取得しました: $destPath" -ForegroundColor Green
        } else {
            Write-Host "❌ chromedriver.exe がzip内に見つかりません" -ForegroundColor Red
            exit 1
        }
        Remove-Item $tmpZip -Force
        Remove-Item $tmpExtract -Recurse -Force
    } catch {
        Write-Host "❌ ChromeDriverのダウンロードに失敗: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Chrome/ChromeDriverバージョンチェック
$chromeVer = Get-ChromeVersion -ExplicitPath $ChromeBinaryPath
if (-not $chromeVer) {
    if ($ChromeBinaryPath) {
        Write-Host "❌ 指定された ChromeBinaryPath が無効、またはバージョン取得不可: $ChromeBinaryPath" -ForegroundColor Red
    } else {
        Write-Host "❌ Google Chromeが見つかりません。-ChromeBinaryPath でパスを指定するか Chrome をインストールしてください。" -ForegroundColor Red
    }
    exit 1
}
$chromeMajor = $chromeVer.Split('.')[0]

$driverVer = Get-ChromeDriverVersion $chromeDriverExe
$needDownload = $false
if (-not $driverVer) {
    $needDownload = $true
    Write-Host "⚠️  ChromeDriverが見つかりません。自動取得します。" -ForegroundColor Yellow
} elseif ($driverVer.Split('.')[0] -ne $chromeMajor) {
    $needDownload = $true
    Write-Host "⚠️  ChromeDriverバージョン($driverVer)とChrome($chromeVer)が一致しません。自動更新します。" -ForegroundColor Yellow
}
if ($needDownload) {
    # Chrome for Testing APIで最新バージョンを取得
    $verApi = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
    try {
        $json = Invoke-RestMethod -Uri $verApi -UseBasicParsing
        $ver = $json.channels.Stable.version
        if ($ver.Split('.')[0] -eq $chromeMajor) {
            Download-ChromeDriver -version $ver -destPath $chromeDriverExe
        } else {
            Write-Host "❌ Chromeのメジャーバージョン($chromeMajor)に対応するChromeDriverが見つかりません。" -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "❌ ChromeDriverバージョン情報の取得に失敗: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host "🎯 BOOTH CSV 自動スクリーンショット撮影開始" -ForegroundColor Green
Write-Host "💻 PowerShell バージョン: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
Write-Host "📁 出力先: $OutputDir" -ForegroundColor Cyan
Write-Host "🌐 HTMLファイル: $HtmlFile" -ForegroundColor Cyan
Write-Host ""

$driver = $null

try {
    # Chrome起動（現在のフォルダのChromeDriverを使用）
    Write-Host "🚀 Chrome起動中..." -ForegroundColor Yellow
    
    # ChromeDriverService作成
    $chromeDriverService = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService((Get-Location).Path)
    
    # ChromeOptionsでログレベルを設定
    $chromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $chromeOptions.AddArgument("--enable-logging")
    $chromeOptions.AddArgument("--log-level=0")
    $chromeOptions.SetLoggingPreference([OpenQA.Selenium.LogType]::Browser, [OpenQA.Selenium.LogLevel]::All)
    
    # ChromeDriverを手動で作成
    if ($ChromeBinaryPath) {
        Write-Host "ℹ️  指定Chromeバイナリを使用: $ChromeBinaryPath" -ForegroundColor Cyan
        $chromeOptions.BinaryLocation = $ChromeBinaryPath
    }
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($chromeDriverService, $chromeOptions)
    $driver.Manage().Window.Size = [System.Drawing.Size]::new(1200, 1000)
    
    # HTTPサーバー経由でHTMLファイルを開く
    $serverUrl = "http://localhost:8080"
    Write-Host "🌐 HTTPサーバー経由でアクセス: $serverUrl" -ForegroundColor Cyan
    $driver.Url = $serverUrl
    Start-Sleep 8  # サーバー起動とページロード、JavaScript初期化を十分に待つ
    
    # DOM要素の読み込み確認
    $driver.ExecuteScript(@"
        console.log('=== DOM要素確認 ===');
        console.log('file要素:', document.getElementById('file'));
        console.log('labelyn要素:', document.getElementById('labelyn'));
        console.log('customLabelEnable要素:', document.getElementById('customLabelEnable'));
        console.log('orderImageEnable要素:', document.getElementById('orderImageEnable'));
        console.log('DOM確認完了');
"@)
    
    Write-Host "✅ ページロード完了" -ForegroundColor Green
    
    # コンソールログ取得関数を定義
    function Get-BrowserConsoleLog {
        param($driver)
        try {
            # JavaScriptでコンソールログを直接取得する方法
            $jsScript = @"
                // コンソールログを配列に保存する仕組みを実装
                if (!window.capturedLogs) {
                    window.capturedLogs = [];
                    
                    // console.logを拡張してログを保存
                    const originalLog = console.log;
                    const originalError = console.error;
                    const originalWarn = console.warn;
                    const originalInfo = console.info;
                    
                    console.log = function(...args) {
                        window.capturedLogs.push({level: 'LOG', message: args.join(' '), timestamp: new Date().toISOString()});
                        originalLog.apply(console, args);
                    };
                    
                    console.error = function(...args) {
                        window.capturedLogs.push({level: 'ERROR', message: args.join(' '), timestamp: new Date().toISOString()});
                        originalError.apply(console, args);
                    };
                    
                    console.warn = function(...args) {
                        window.capturedLogs.push({level: 'WARN', message: args.join(' '), timestamp: new Date().toISOString()});
                        originalWarn.apply(console, args);
                    };
                    
                    console.info = function(...args) {
                        window.capturedLogs.push({level: 'INFO', message: args.join(' '), timestamp: new Date().toISOString()});
                        originalInfo.apply(console, args);
                    };
                }
                
                // 最新のログを取得して返す
                const recentLogs = window.capturedLogs.slice(-20); // 最新20件
                window.capturedLogs = []; // ログをクリア
                return JSON.stringify(recentLogs);
"@
            
            $logsJson = $driver.ExecuteScript($jsScript)
            if ($logsJson -and $logsJson -ne "[]") {
                $logs = $logsJson | ConvertFrom-Json
                if ($logs.Count -gt 0) {
                    Write-Host "📜 ブラウザコンソールログ:" -ForegroundColor Blue
                    foreach ($log in $logs) {
                        $timestamp = ([DateTime]$log.timestamp).ToString("HH:mm:ss.fff")
                        $level = $log.level
                        $message = $log.message
                        
                        $color = switch ($level) {
                            "ERROR" { "Red" }
                            "WARN" { "Yellow" }
                            "INFO" { "Cyan" }
                            "LOG" { "White" }
                            default { "White" }
                        }
                        
                        Write-Host "  [$timestamp] $level`: $message" -ForegroundColor $color
                    }
                    Write-Host ""
                }
            }
        } catch {
            Write-Host "⚠️ コンソールログ取得エラー: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # 初期化時のコンソールログを取得
    Write-Host "📋 初期化時のコンソールログ:" -ForegroundColor Blue
    Get-BrowserConsoleLog -driver $driver
    
    Write-Host ""
    
    # スクリーンショット撮影リスト
    $screenshots = @(
        @{
            Name = "main-interface.png"
            Description = "メインインターフェース（初期状態）"
            Height = 600
            Script = @"
                // DOM要素の存在確認
                const fileElement = document.getElementById('file');
                const labelynElement = document.getElementById('labelyn');
                const customLabelElement = document.getElementById('customLabelEnable');
                const orderImageElement = document.getElementById('orderImageEnable');
                
                console.log('DOM要素確認:');
                console.log('file:', fileElement);
                console.log('labelyn:', labelynElement);
                console.log('customLabelEnable:', customLabelElement);
                console.log('orderImageEnable:', orderImageElement);
                
                if (!fileElement || !labelynElement || !customLabelElement || !orderImageElement) {
                    console.error('❌ 必要なDOM要素が見つかりません');
                    return;
                }
                
                // 初期状態にリセット
                fileElement.value = '';
                labelynElement.checked = true;
                customLabelElement.checked = false;
                orderImageElement.checked = false;
                
                // チェックボックス変更イベントを発火
                labelynElement.dispatchEvent(new Event('change'));
                customLabelElement.dispatchEvent(new Event('change'));
                orderImageElement.dispatchEvent(new Event('change'));
                
                // ラベルスキップ面数を0に設定（初期状態）
                const skipCountInput = document.getElementById('labelskipnum');
                if (skipCountInput) {
                    skipCountInput.value = '0';
                    skipCountInput.dispatchEvent(new Event('change'));
                    console.log('✅ ラベルスキップ面数を0に設定（初期状態）');
                }
                
                // ページトップにスクロール
                window.scrollTo(0, 0);
                
                console.log('✅ 初期状態設定完了');
"@
            PostScript = @"
                // 撮影後にラベルスキップ面数を3に設定（機能紹介のため）
                const skipCountInput = document.getElementById('labelskipnum');
                if (skipCountInput) {
                    skipCountInput.value = '3';
                    skipCountInput.dispatchEvent(new Event('change'));
                    console.log('✅ ラベルスキップ面数を3に設定（機能紹介のため）');
                }
                console.log('✅ 初期状態スクリーンショット撮影完了');
"@
            Wait = 5
        },
        @{
            Name = "custom-labels.png"
            Description = "カスタムラベル機能（設定画面）"
            Height = 1200
            Script = @"
                // DOM要素の存在確認
                const labelynElement = document.getElementById('labelyn');
                const customLabelElement = document.getElementById('customLabelEnable');
                const orderImageElement = document.getElementById('orderImageEnable');
                window.scrollCompleted = false;
                if (!labelynElement || !customLabelElement || !orderImageElement) {
                    console.error('❌ 必要なDOM要素が見つかりません');
                    window.scrollCompleted = true;
                    return;
                }
                // カスタムラベル機能のみON
                labelynElement.checked = true;
                customLabelElement.checked = true;
                orderImageElement.checked = false;
                // 各要素のchangeイベントを発火
                labelynElement.dispatchEvent(new Event('change'));
                customLabelElement.dispatchEvent(new Event('change'));
                orderImageElement.dispatchEvent(new Event('change'));
                // 印刷面数を5に設定
                const printCountInput = document.querySelector('.custom-label-count-group input');
                if (printCountInput) {
                    printCountInput.value = '5';
                    printCountInput.dispatchEvent(new Event('input'));
                    console.log('✅ 印刷面数を5に設定');
                }
                // 「残りラベルに任意文字列を印刷」まで即時スクロール
                const customLabelRow = document.getElementById('customLabelRow');
                if (customLabelRow) {
                    // 「残りラベルに任意文字列を印刷」のラベル部分を探す
                    const customLabelLabel = document.querySelector('label[for="customLabelEnable"]');
                    if (customLabelLabel) {
                        customLabelLabel.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        window.scrollBy(0, -80); // 固定ヘッダー分を調整
                        console.log('✅ 「残りラベルに任意文字列を印刷」までスクロール完了');
                        window.scrollCompleted = true;
                    } else {
                        // フォールバック: customLabelRowにスクロール
                        customLabelRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        window.scrollBy(0, -80); // 固定ヘッダー分を調整
                        console.log('✅ カスタムラベル設定エリアにスクロール完了（フォールバック）');
                        window.scrollCompleted = true;
                    }
                } else {
                    console.log('⚠️ カスタムラベル設定エリアが見つかりません');
                    window.scrollCompleted = true;
                }
                console.log('✅ カスタムラベル設定完了');
"@
            PostScript = @"
                // 最初から存在する.rich-text-editorにサンプルテキストを入力
                setTimeout(() => {
                    const editors = document.querySelectorAll('.rich-text-editor');
                    if (editors.length > 0) {
                        editors[0].innerHTML = '<div><b>【サンプル商品】</b></div><div>アクリルキーホルダー</div><div style=\"color: #666;\">商品コード: AKH-001</div>';
                        editors[0].dispatchEvent(new Event('input'));
                        console.log('✅ サンプルテキスト入力完了');

                        // カスタムフォント設定セクションを展開
                        setTimeout(() => {
                            // フォントセクションが折りたたまれている場合は展開
                            const fontSectionContent = document.getElementById('fontSectionContent');
                            if (fontSectionContent && fontSectionContent.style.maxHeight === '0px') {
                                if (typeof toggleFontSection === 'function') {
                                    toggleFontSection();
                                    console.log('✅ カスタムフォント設定セクションを展開');
                                } else {
                                    console.log('⚠️ toggleFontSection関数が見つかりません');
                                }
                            } else {
                                console.log('✅ カスタムフォント設定セクションは既に展開済み');
                            }
                        }, 500);

                        // 「【サンプル商品】」テキストを選択
                        setTimeout(() => {
                            const range = document.createRange();
                            const selection = window.getSelection();

                            // <b>タグ内のテキストノードを探す
                            const boldElement = editors[0].querySelector('b');
                            if (boldElement && boldElement.firstChild) {
                                range.selectNodeContents(boldElement);
                                selection.removeAllRanges();
                                selection.addRange(range);
                                console.log('✅ 「【サンプル商品】」テキスト選択完了');

                                // 右クリック前に200px下にスクロール
                                window.scrollBy(0, 200);
                                console.log('✅ 右クリック前に200px下にスクロールしました');

                                // 右クリックイベントを発生させる
                                setTimeout(() => {
                                    const contextMenuEvent = new MouseEvent('contextmenu', {
                                        bubbles: true,
                                        cancelable: true,
                                        view: window,
                                        button: 2,
                                        clientX: boldElement.getBoundingClientRect().left + 50,
                                        clientY: boldElement.getBoundingClientRect().top + 10
                                    });
                                    boldElement.dispatchEvent(contextMenuEvent);
                                    console.log('✅ 右クリックメニュー表示');
                                }, 500);
                            }
                        }, 1000);
                    }

                    // カスタムラベル設定エリアにスクロール
                    const customRow = document.getElementById('customLabelRow');
                    if (customRow) {
                        customRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        // 固定ヘッダー分を考慮して調整
                        setTimeout(() => {
                            window.scrollBy(0, -80); // 固定ヘッダー分を調整
                        }, 1000);
                    }
                }, 500);
"@
            AdditionalShots = @(
                @{
                    Name = "custom-labels-sheet.png"
                    Description = "カスタムラベル機能（シート表示位置）"
                    ScrollScript = @"
                        // csv-labels.pngと同じ位置にスクロール（section.sheet table.label44）
                        const firstLabelTable = document.querySelector('section.sheet table.label44');
                        if (firstLabelTable) {
                            firstLabelTable.scrollIntoView({ behavior: 'smooth', block: 'start' });
                            // 固定ヘッダー分を考慮して少し上にスクロール調整
                            setTimeout(() => {
                                window.scrollBy(0, -80); // 固定ヘッダー分を調整
                                console.log('✅ カスタムラベルシート位置(section.sheet table.label44)にスクロール＋ヘッダー調整完了');
                            }, 1000);
                        } else {
                            console.log('⚠️ ラベルテーブル(section.sheet table.label44)が見つかりません - ページ下部にスクロールします');
                            window.scrollTo(0, document.body.scrollHeight);
                        }
"@
                },
                @{
                    Name = "custom-fonts.png"
                    Description = "カスタムフォント設定セクション"
                    Height = 600
                    ScrollScript = @"
                        // カスタムフォント設定セクションにスクロール
                        const fontSection = document.getElementById('fontSectionContent');
                        if (fontSection) {
                            fontSection.scrollIntoView({ behavior: 'smooth', block: 'center' });
                            // 固定ヘッダー分を考慮して調整
                            setTimeout(() => {
                                window.scrollBy(0, -80); // 固定ヘッダー分を調整
                            }, 1000);
                            console.log('✅ カスタムフォント設定セクションにスクロール完了');
                        } else {
                            // フォント設定セクションが見つからない場合はカスタムラベル設定エリアにスクロール
                            const customRow = document.getElementById('customLabelRow');
                            if (customRow) {
                                customRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                                // フォント設定エリアまで少し下にスクロール + 固定ヘッダー調整
                                setTimeout(() => {
                                    window.scrollBy(0, 200 - 80); // フォント設定エリア移動 + 固定ヘッダー分を調整
                                    console.log('✅ カスタムラベル設定エリア下部（フォント設定付近）にスクロール完了');
                                }, 1000);
                            } else {
                                console.log('⚠️ カスタムラベル設定エリアが見つかりません');
                            }
                        }
"@
                }
            )
            Wait = 2
        },
        @{
            Name = "image-function.png"
            Description = "画像表示機能（チェックボックスON状態）"
            Height = 600
            Script = @"
                // DOM要素の存在確認
                const labelynElement = document.getElementById('labelyn');
                const customLabelElement = document.getElementById('customLabelEnable');
                const orderImageElement = document.getElementById('orderImageEnable');
                
                if (!labelynElement || !customLabelElement || !orderImageElement) {
                    console.error('❌ 必要なDOM要素が見つかりません');
                    return;
                }
                
                // 画像機能とカスタムラベル機能をON（ドロップゾーンは空の状態）
                labelynElement.checked = true;
                customLabelElement.checked = true;
                orderImageElement.checked = true;
                
                // 各要素のchangeイベントを発火
                labelynElement.dispatchEvent(new Event('change'));
                customLabelElement.dispatchEvent(new Event('change'));
                orderImageElement.dispatchEvent(new Event('change'));
                
                // 画像ドロップゾーンエリアにスクロール
                setTimeout(() => {
                    const imageRow = document.getElementById('orderImageRow');
                    if (imageRow) {
                        imageRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        // 固定ヘッダー分を考慮して調整
                        setTimeout(() => {
                            window.scrollBy(0, -80); // 固定ヘッダー分を調整
                        }, 1000);
                    }
                }, 1000);
                
                console.log('✅ 画像機能設定完了');
"@
            AdditionalShots = @(
                @{
                    Name = "image-dropzone.png"
                    Description = "画像ドロップゾーン（画像アップロード後）"
                    Height = 600
                    ScrollScript = @"
                        // 画像機能をONにして、実際に画像をアップロード状態にする
                        console.log('✅ 画像機能ON完了');
                        
                        // アプリケーションの初期化を待ってから画像をアップロード
                        setTimeout(() => {
                            fetch('http://localhost:8080/sample/footersample.png')
                                .then(response => response.blob())
                                .then(async blob => {
                                    const file = new File([blob], 'footersample.png', { type: 'image/png' });
                                    const arrayBuffer = await blob.arrayBuffer();
                                    const blobUrl = URL.createObjectURL(new Blob([arrayBuffer], { type: blob.type }));

                                    // orderImageDropZone 初期化待ち（最大5秒リトライ）
                                    const start = performance.now();
                                    (function trySet(){
                                        if (window.orderImageDropZone && window.orderImageDropZone.setImage) {
                                            window.orderImageDropZone.setImage(blobUrl, arrayBuffer);
                                            console.log('グローバル画像を設定しました(Blob URL + ArrayBuffer):', file.name);

                                            const imageRow = document.getElementById('orderImageRow');
                                            if (imageRow) {
                                                imageRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                                                setTimeout(() => { window.scrollBy(0, -80); }, 1000);
                                            }
                                        } else if (performance.now() - start < 5000) {
                                            setTimeout(trySet, 200);
                                        } else {
                                            console.warn('orderImageDropZoneが初期化されず画像設定を断念しました');
                                        }
                                    })();
                                })
                                .catch(error => console.log('画像読み込みエラー:', error));
                        }, 3000);
"@
                }
            )
            Wait = 5
        },
        @{
            Name = "csv-labels.png"
            Description = "CSV読み込み後のラベル印刷プレビュー"
            Height = 1000
            Script = @"
                console.log('=== CSV+画像読み込み（ラベル用）開始 ===');
                
                // DOM要素の存在確認
                const labelynElement = document.getElementById('labelyn');
                const customLabelElement = document.getElementById('customLabelEnable');
                const orderImageElement = document.getElementById('orderImageEnable');
                const fileElement = document.getElementById('file');
                
                if (!labelynElement || !customLabelElement || !orderImageElement || !fileElement) {
                    console.error('❌ 必要なDOM要素が見つかりません');
                    return;
                }
                
                // 1. 全ての機能をONにする
                labelynElement.checked = true;
                customLabelElement.checked = true;
                orderImageElement.checked = true;
                
                // 各要素のchangeイベントを発火
                labelynElement.dispatchEvent(new Event('change'));
                customLabelElement.dispatchEvent(new Event('change'));
                orderImageElement.dispatchEvent(new Event('change'));
                console.log('✅ 全機能ON完了');
                
                // 2. 画像を設定
                fetch('http://localhost:8080/sample/footersample.png')
                    .then(response => response.blob())
                    .then(async blob => {
                        const arrayBuffer = await blob.arrayBuffer();
                        const blobUrl = URL.createObjectURL(new Blob([arrayBuffer], { type: blob.type }));
                        const start = performance.now();
                        (function trySet(){
                            if (window.orderImageDropZone && window.orderImageDropZone.setImage) {
                                window.orderImageDropZone.setImage(blobUrl, arrayBuffer);
                                console.log('✅ 画像設定完了 (Blob URL + ArrayBuffer)');
                            } else if (performance.now() - start < 5000) {
                                setTimeout(trySet, 200);
                            } else {
                                console.warn('画像設定失敗: orderImageDropZone未初期化');
                            }
                        })();
                    });
                
                // 3. CSVファイルを読み込み（QRドロップゾーンより先に実行）
                setTimeout(() => {
                    console.log('=== CSV読み込み開始 ===');
                    fetch('http://localhost:8080/sample/booth_orders_sample.csv')
                        .then(response => response.blob())
                        .then(blob => {
                            const file = new File([blob], 'booth_orders_sample.csv', { type: 'text/csv' });
                            const fileInput = document.getElementById('file');
                            
                            // DataTransferを使ってファイルを設定
                            const dt = new DataTransfer();
                            dt.items.add(file);
                            Object.defineProperty(fileInput, 'files', {
                                value: dt.files,
                                writable: false,
                            });
                            
                            console.log('✅ CSVファイル設定完了:', fileInput.files.length);
                            
                            // changeイベントを発火
                            const changeEvent = new Event('change', { bubbles: true, cancelable: true });
                            fileInput.dispatchEvent(changeEvent);
                            console.log('🔥 changeイベント発火');
                            
                            // autoProcessCSV()を直接呼び出してカスタムラベルも含めて処理（非同期処理として実行）
                            setTimeout(async () => {
                                if (typeof autoProcessCSV === 'function') {
                                    console.log('🔄 autoProcessCSV()を直接実行');
                                    await autoProcessCSV();
                                    console.log('✅ autoProcessCSV()実行完了');
                                } else {
                                    console.log('⚠️ autoProcessCSV関数が見つかりません');
                                }
                            }, 1000);
                            
                            // CSV処理完了後にQRコード機能をデモ
                            setTimeout(() => {
                                console.log('=== QRコード機能デモ開始 ===');
                                const qrDropZones = document.querySelectorAll('.dropzone');
                                if (qrDropZones.length > 0) {
                                    console.log('✅ QRドロップゾーンを発見:', qrDropZones.length + '個');
                                    
                                    // 最初のQRドロップゾーンにサンプル画像を設定
                                    const firstDropZone = qrDropZones[0];
                                    
                                    fetch('http://localhost:8080/sample/qrcodedsample.png')
                                        .then(response => response.blob())
                                        .then(blob => {
                                            // 直接的なファイル入力をシミュレート
                                            const reader = new FileReader();
                                            reader.onload = function(e) {
                                                try {
                                                    // QRドロップゾーンに直接画像を追加
                                                    const elImage = document.createElement('img');
                                                    elImage.src = e.target.result;
                                                    elImage.style.maxWidth = '100%';
                                                    elImage.style.height = 'auto';
                                                    
                                                    // 画像読み込み完了後にQR処理を実行
                                                    elImage.onload = function() {
                                                        // ドロップゾーンを非表示にして画像を表示
                                                        firstDropZone.style.display = 'none';
                                                        firstDropZone.parentNode.appendChild(elImage);
                                                        
                                                        // QR読み取り処理を実行（もし利用可能であれば）
                                                        if (typeof readQR === 'function') {
                                                            readQR(elImage);
                                                        }
                                                        
                                                        // リセットイベントを追加（もし利用可能であれば）
                                                        if (typeof addEventQrReset === 'function') {
                                                            addEventQrReset(elImage);
                                                        }
                                                        
                                                        console.log('✅ QRコード画像を直接設定完了');
                                                    };
                                                } catch (error) {
                                                    console.log('QRコード設定エラー:', error);
                                                }
                                            };
                                            reader.readAsDataURL(blob);
                                        })
                                        .catch(error => console.log('QRコード画像読み込みエラー:', error));
                                } else {
                                    console.log('⚠️ QRドロップゾーンが見つかりません');
                                }
                            }, 5000); // CSV処理完了から5秒後にQRコード設定
                        })
                        .catch(error => console.error('❌ CSV読み込みエラー:', error));
                }, 2000);
"@
            PostScript = @"
                // 処理完了まで待機してラベルセクションにスクロール（同期的に実行）
                console.log('=== ラベルセクション処理完了待機開始 ===');
                
                const maxRetries = 15;  // リトライ回数を増加
                let retryCount = 0;
                
                const checkAndScrollToLabels = () => {
                    retryCount++;
                    console.log('ラベル処理状況チェック試行:', retryCount);
                    
                    // 正しいDOM構造で注文明細を探す
                    const sheetElements = document.querySelectorAll('.sheet');
                    const label44Elements = document.querySelectorAll('.sheet .label44');
                    const pageElements = document.querySelectorAll('.sheet .page');
                    const printDisplay = document.getElementById('printCountDisplay');
                    const fileInput = document.getElementById('file');
                    
                    console.log({
                        'Sheet数': sheetElements.length,
                        'Label44数': label44Elements.length,
                        'Page数': pageElements.length,
                        'ファイル数': fileInput.files.length,
                        '印刷表示': printDisplay ? printDisplay.style.display : 'なし'
                    });
                    
                    // label44要素またはpage要素があればCSV処理完了とみなす
                    if (label44Elements.length > 0 || pageElements.length > 0) {
                        console.log('✅ CSV処理完了！ラベルセクションを表示');
                        
                        // 最初のsection.sheetのtable.label44までスクロール
                        const firstLabelTable = document.querySelector('section.sheet table.label44');
                        if (firstLabelTable) {
                            firstLabelTable.scrollIntoView({ behavior: 'smooth', block: 'start' });
                            console.log('🔄 スクロール実行中...');
                            // スクロール完了を待機
                            setTimeout(() => {
                                window.scrollBy(0, -80); // 固定ヘッダー分を調整
                                console.log('✅ ラベルテーブル(section.sheet table.label44)にスクロール＋ヘッダー調整完了');
                                // スクロール完了マーカーを設定
                                window.scrollCompleted = true;
                            }, 2000);
                        } else {
                            console.log('⚠️ ラベルテーブル(section.sheet table.label44)が見つかりません');
                            window.scrollCompleted = true;
                        }
                        return true;
                    } else if (retryCount < maxRetries) {
                        console.log('⏳ 処理中...リトライします');
                        setTimeout(checkAndScrollToLabels, 1500);  // リトライ間隔を短縮
                        return false;
                    } else {
                        console.log('❌ タイムアウト：注文明細が生成されませんでした');
                        window.scrollCompleted = true;
                        return false;
                    }
                };
                
                // 初回チェックを1秒後に開始
                setTimeout(checkAndScrollToLabels, 1000);
"@
            AdditionalShots = @(
                @{
                    Name = "csv-orders.png"
                    Description = "CSV読み込み後の注文明細印刷プレビュー"
                    Height = 1300
                    ScrollScript = @"
                        // CSV読み込み済みのため、単純に注文明細セクションにスクロール
                        const firstPageDiv = document.querySelector('section.sheet div.page');
                        if (firstPageDiv) {
                            firstPageDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
                            // 固定ヘッダー分を考慮して少し上にスクロール調整
                            setTimeout(() => {
                                window.scrollBy(0, -80); // 固定ヘッダー分を調整
                                console.log('✅ 注文明細ページ(section.sheet div.page)にスクロール完了');
                            }, 1000);
                        } else {
                            console.log('⚠️ 注文明細ページ(section.sheet div.page)が見つかりません');
                        }
"@
                }
            )
            Wait = 5
        },
        @{
            Name = "usage-guide.png"
            Description = "全機能有効状態（使用方法ガイド）"
            Height = 1600
            Script = @"
                // 全ての機能をONにした状態
                document.getElementById('labelyn').checked = true;
                document.getElementById('customLabelEnable').checked = true;
                document.getElementById('orderImageEnable').checked = true;
                
                // 各機能をアクティブ化（イベント発火）
                document.getElementById('labelyn').dispatchEvent(new Event('change'));
                document.getElementById('customLabelEnable').dispatchEvent(new Event('change'));
                document.getElementById('orderImageEnable').dispatchEvent(new Event('change'));
                
                // ページトップに戻って全体を表示
                window.scrollTo(0, 0);
                
                // ブラウザウィンドウサイズを縦に延ばして全機能が画面内におさまるようにする
                console.log('ウィンドウサイズを調整します');
"@
            PostScript = @"
                // PowerShellからウィンドウサイズを変更するため、一旦待機
                console.log('ウィンドウサイズ調整待機');
"@
            Wait = 5
        }
    )
    
    # 撮影実行
    $successCount = 0
    
    foreach ($shot in $screenshots) {
        Write-Host "📸 撮影中: $($shot.Description) ($($shot.Name))" -ForegroundColor Cyan
        
        try {
            # 各スクリーンショットの高さ設定に基づいてウィンドウサイズを調整
            $windowHeight = if ($shot.Height) { $shot.Height } else { 1000 }
            Write-Host "📏 ウィンドウサイズを調整中... (1200 x $windowHeight)" -ForegroundColor Yellow
            $driver.Manage().Window.Size = [System.Drawing.Size]::new(1200, $windowHeight)
            Start-Sleep 1
            
            # メインスクリプト実行
            $driver.ExecuteScript($shot.Script)
            
            # 待機時間（カスタム設定があれば使用、なければ2秒）
            $waitTime = if ($shot.Wait) { $shot.Wait } else { 2 }
            Start-Sleep $waitTime
            
            # ポストスクリプトがあれば実行
            if ($shot.PostScript) {
                $driver.ExecuteScript($shot.PostScript)
                # スクロール完了を確認（csv-labels.png, custom-labels.png）
                if ($shot.Name -eq "csv-labels.png" -or $shot.Name -eq "custom-labels.png") {
                    $scrollCompleted = $false
                    $scrollRetries = 0
                    $maxScrollRetries = 30  # 最大30秒待機
                    Write-Host "⏳ スクロール完了を待機中..." -ForegroundColor Yellow
                    while (-not $scrollCompleted -and $scrollRetries -lt $maxScrollRetries) {
                        Start-Sleep 1
                        $scrollRetries++
                        try {
                            $scrollStatus = $driver.ExecuteScript("return window.scrollCompleted || false;")
                            if ($scrollStatus -eq $true) {
                                $scrollCompleted = $true
                                Write-Host "✅ スクロール完了確認 ($scrollRetries 秒)" -ForegroundColor Green
                            }
                        } catch {
                            # スクロール状態取得エラーは無視
                        }
                    }
                    if (-not $scrollCompleted) {
                        Write-Host "⚠️ スクロール完了タイムアウト" -ForegroundColor Yellow
                    }
                } else {
                    Start-Sleep 2  # 通常のPostScript完了待機
                }
            }
            
            # スクリーンショット撮影前の固定待機
            Start-Sleep 2
            
            # スクリーンショット撮影
            $outputPath = Join-Path $OutputDir $shot.Name
            $driver.GetScreenshot().SaveAsFile($outputPath, "Png")
            Write-Host "✅ 保存完了: $outputPath" -ForegroundColor Green
            $successCount++
            
            # コンソールログを取得・表示
            Get-BrowserConsoleLog -driver $driver
            
            # custom-labels.pngの場合は右クリックメニューを閉じる
            if ($shot.Name -eq "custom-labels.png") {
                Write-Host "🔧 右クリックメニューを閉じています..." -ForegroundColor Yellow
                $driver.ExecuteScript(@"
                    // 右クリックメニューを閉じる
                    document.addEventListener('click', function() {}, { once: true });
                    document.body.click();
                    
                    // 選択も解除する
                    const selection = window.getSelection();
                    if (selection) {
                        selection.removeAllRanges();
                    }
                    
                    console.log('✅ 右クリックメニューと選択を解除しました');
"@)
                Start-Sleep 1
            }
            
            # AdditionalShotsがある場合は追加撮影（複数対応）
            if ($shot.AdditionalShots) {
                for ($i = 0; $i -lt $shot.AdditionalShots.Length; $i++) {
                    $additionalShot = $shot.AdditionalShots[$i]
                    $shotNumber = $i + 1
                    Write-Host "📸 追加撮影中 ($shotNumber/$($shot.AdditionalShots.Length)): $($additionalShot.Description) ($($additionalShot.Name))" -ForegroundColor Magenta
                    
                    try {
                        # AdditionalShotに高さ設定がある場合はウィンドウサイズを調整
                        if ($additionalShot.Height) {
                            Write-Host "📏 追加撮影用ウィンドウサイズを調整中... (1200 x $($additionalShot.Height))" -ForegroundColor Yellow
                            $driver.Manage().Window.Size = [System.Drawing.Size]::new(1200, $additionalShot.Height)
                            Start-Sleep 1
                        }
                        
                        # スクロールスクリプト実行
                        if ($additionalShot.ScrollScript) {
                            $driver.ExecuteScript($additionalShot.ScrollScript)
                            Start-Sleep 3  # スクロール完了まで待機
                        }
                        
                        # 追加スクリーンショット撮影前の固定待機
                        Start-Sleep 2
                        
                        # 追加スクリーンショット撮影
                        $additionalOutputPath = Join-Path $OutputDir $additionalShot.Name
                        $driver.GetScreenshot().SaveAsFile($additionalOutputPath, "Png")
                        Write-Host "✅ 追加保存完了 ($shotNumber/$($shot.AdditionalShots.Length)): $additionalOutputPath" -ForegroundColor Green
                        $successCount++
                        
                        # 追加撮影でもコンソールログを取得
                        Get-BrowserConsoleLog -driver $driver
                        
                    } catch {
                        Write-Host "❌ 追加撮影エラー ($shotNumber/$($shot.AdditionalShots.Length)) ($($additionalShot.Name)): $($_.Exception.Message)" -ForegroundColor Red
                        # 追加撮影エラー時もコンソールログを確認
                        Get-BrowserConsoleLog -driver $driver
                    }
                }
            }
            
        } catch {
            Write-Host "❌ エラー ($($shot.Name)): $($_.Exception.Message)" -ForegroundColor Red
            # エラー時もコンソールログを確認
            Get-BrowserConsoleLog -driver $driver
        }
        
        Start-Sleep 1
    }
    
    Write-Host ""
    Write-Host "📊 撮影結果: $successCount/$($screenshots.Count) 件成功" -ForegroundColor Yellow
    
} catch {
    Write-Host "❌ 実行エラー: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "スタックトレース:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
} finally {
    if ($driver) {
        $driver.Quit()
        Write-Host "🔧 ブラウザ終了" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "🎉 撮影完了！" -ForegroundColor Green
Write-Host "📁 出力ディレクトリ: $OutputDir" -ForegroundColor Cyan
Write-Host "確認コマンド: ls $OutputDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "📝 コミット用コマンド:" -ForegroundColor Yellow
Write-Host "cd ../..; git add docs/images/; git commit -m 'docs: PowerShell 7による自動スクリーンショット更新'" -ForegroundColor White
