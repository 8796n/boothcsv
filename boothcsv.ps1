<#
boothcsv.ps1
簡易HTTPファイルサーバー (PowerShell)
- カレントディレクトリを配信
- ルートにアクセスが来た場合は `boothcsv.html` を返す
- 使い方: PowerShell を管理者権限不要で開き、スクリプトのあるフォルダで実行
#>
[CmdletBinding()]
param(
    [int]$Port = 8796,
    [string]$DefaultPage = "boothcsv.html",
    [bool]$OpenBrowser = $true
)
Add-Type -AssemblyName System.Web

$listener = New-Object System.Net.HttpListener
$prefix = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)
try {
    $listener.Start()
} catch {
    Write-Error "HTTP Listener を開始できませんでした: $_"
    exit 1
}

Write-Host "Serving $PWD at $prefix  (Ctrl+C で停止)"

# 起動時に既定のブラウザでルート URL を開く
if ($OpenBrowser) {
    try {
        Start-Process $prefix -ErrorAction Stop
    } catch {
        # 万が一 Start-Process で直接開けない環境向けに explorer.exe を試す
        try {
            Start-Process "explorer.exe" $prefix -ErrorAction Stop
        } catch {
            Write-Warning "ブラウザを自動起動できませんでした: $_"
        }
    }
}

function Get-ContentType($ext){
    switch ($ext.ToLower()) {
        ".html" { "text/html; charset=utf-8"; break }
        ".htm"  { "text/html; charset=utf-8"; break }
        ".js"   { "application/javascript; charset=utf-8"; break }
        ".css"  { "text/css; charset=utf-8"; break }
        ".json" { "application/json; charset=utf-8"; break }
        ".png"  { "image/png"; break }
        ".jpg"  { "image/jpeg"; break }
        ".jpeg" { "image/jpeg"; break }
        ".gif"  { "image/gif"; break }
        ".svg"  { "image/svg+xml; charset=utf-8"; break }
        ".txt"  { "text/plain; charset=utf-8"; break }
        default { "application/octet-stream" }
    }
}

try {
    # Ctrl+C を検知したら listener.Stop() を呼んで GetContext のブロッキングを解除する
    $null = [Console]::add_CancelKeyPress([ConsoleCancelEventHandler]{
        param($sender,$e)
        Write-Host "Ctrl+C detected - stopping listener..."
        try { $listener.Stop() } catch {}
        # スクリプトは終了処理を行うので、Ctrl+C の既定動作は抑制する
        $e.Cancel = $true
    })

    while ($listener.IsListening) {
        try {
            $ctx = $listener.GetContext()
        } catch [System.Net.HttpListenerException] {
            # listener が Stop() されたときなどに発生する。ループを抜けて終了処理へ。
            break
        } catch [System.InvalidOperationException] {
            # こちらも listener の状態変化で発生しうるので安全に抜ける
            break
        }

        $req = $ctx.Request
        $res = $ctx.Response

    # アクセスログ（デフォルトでコンソールに表示）
        try {
            $remote = if ($req.RemoteEndPoint -ne $null) { $req.RemoteEndPoint.Address.ToString() } elseif ($req.UserHostAddress) { $req.UserHostAddress } else { 'unknown' }
        } catch {
            $remote = 'unknown'
        }
    Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] $($req.HttpMethod) $($req.Url.AbsolutePath) from $remote"

        # パスを安全に処理
        $rawPath = $req.Url.AbsolutePath
        $relPath = [System.Web.HttpUtility]::UrlDecode($rawPath.TrimStart('/'))
        if ([string]::IsNullOrEmpty($relPath)) {
            $relPath = $DefaultPage
        }

    # ディレクトリ参照を禁止
        if ($relPath -match "\.\./|\\..") {
            $res.StatusCode = 400
            $msg = "400 Bad Request"
            $buf = [System.Text.Encoding]::UTF8.GetBytes($msg)
            $res.ContentType = "text/plain; charset=utf-8"
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
            $res.OutputStream.Close()
            continue
        }

        # 特別エンドポイント: /proxy?url=... -> サーバー側で外部URLを取得して返す
        if ($relPath -ieq 'proxy') {
            $qs = [System.Web.HttpUtility]::ParseQueryString($req.Url.Query)
            $targetUrl = $qs['url']
            if ([string]::IsNullOrEmpty($targetUrl) -or -not ($targetUrl -match '^https?://')) {
                $res.StatusCode = 400
                $msg = "400 Bad Request - invalid or missing 'url' parameter"
                $buf = [System.Text.Encoding]::UTF8.GetBytes($msg)
                $res.ContentType = "text/plain; charset=utf-8"
                $res.ContentLength64 = $buf.Length
                $res.OutputStream.Write($buf, 0, $buf.Length)
                $res.OutputStream.Close()
                continue
            }

            try {
                $client = New-Object System.Net.Http.HttpClient
                $resp = $client.GetAsync($targetUrl).Result
                if (-not $resp.IsSuccessStatusCode) {
                    $res.StatusCode = 502
                    $msg = "502 Bad Gateway - failed to fetch target"
                    $buf = [System.Text.Encoding]::UTF8.GetBytes($msg)
                    $res.ContentType = "text/plain; charset=utf-8"
                    $res.ContentLength64 = $buf.Length
                    $res.OutputStream.Write($buf, 0, $buf.Length)
                    $res.OutputStream.Close()
                    $client.Dispose()
                    continue
                }

                $bytes = $resp.Content.ReadAsByteArrayAsync().Result
                $ct = $resp.Content.Headers.ContentType
                if ($ct) { $res.ContentType = $ct.ToString() } else { $res.ContentType = "application/octet-stream" }
                $res.ContentLength64 = $bytes.Length
                $res.OutputStream.Write($bytes, 0, $bytes.Length)
                $res.OutputStream.Close()
                $client.Dispose()
                continue
            } catch {
                $res.StatusCode = 500
                $msg = "500 Proxy Error"
                $buf = [System.Text.Encoding]::UTF8.GetBytes($msg)
                $res.ContentType = "text/plain; charset=utf-8"
                $res.ContentLength64 = $buf.Length
                $res.OutputStream.Write($buf, 0, $buf.Length)
                $res.OutputStream.Close()
                continue
            }
        }

        $filePath = Join-Path -Path (Get-Location) -ChildPath $relPath

        if (Test-Path $filePath -PathType Leaf) {
            try {
                $bytes = [System.IO.File]::ReadAllBytes($filePath)
                $res.ContentType = Get-ContentType ([System.IO.Path]::GetExtension($filePath))
                $res.ContentLength64 = $bytes.Length
                $res.OutputStream.Write($bytes, 0, $bytes.Length)
            } catch {
                $res.StatusCode = 500
                $msg = "500 Internal Server Error"
                $buf = [System.Text.Encoding]::UTF8.GetBytes($msg)
                $res.ContentType = "text/plain; charset=utf-8"
                $res.ContentLength64 = $buf.Length
                $res.OutputStream.Write($buf, 0, $buf.Length)
            }
        } else {
            $res.StatusCode = 404
            $msg = "404 Not Found"
            $buf = [System.Text.Encoding]::UTF8.GetBytes($msg)
            $res.ContentType = "text/plain; charset=utf-8"
            $res.ContentLength64 = $buf.Length
            $res.OutputStream.Write($buf, 0, $buf.Length)
        }

        $res.OutputStream.Close()
    }
} finally {
    if ($listener -and $listener.IsListening) {
        $listener.Stop()
        $listener.Close()
    }
}
