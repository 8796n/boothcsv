# シンプルHTTPサーバー for BOOTH CSV Screenshots
param(
    [int]$Port = 8080
)

$projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
Write-Host "🌐 シンプルHTTPサーバー起動"
Write-Host "📁 プロジェクトルート: $projectRoot"
Write-Host "🔗 ポート: $Port"
Write-Host "📖 アクセス: http://localhost:$Port"
Write-Host ""

# HTTPListener作成
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")

try {
    $listener.Start()
    Write-Host "✅ サーバー起動完了: http://localhost:$Port" -ForegroundColor Green
    Write-Host "🛑 停止するには Ctrl+C を押してください" -ForegroundColor Yellow
    Write-Host ""
    
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        $url = $request.Url.AbsolutePath
        
        # デフォルトファイル
        if ($url -eq "/") {
            $url = "/boothcsv.html"
        }
        
        # ファイルパス構築
        $filePath = Join-Path $projectRoot $url.TrimStart('/')
        
        if (Test-Path $filePath) {
            try {
                $content = [System.IO.File]::ReadAllBytes($filePath)
                
                # Content-Type設定
                $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
                switch ($extension) {
                    ".html" { $response.ContentType = "text/html; charset=utf-8" }
                    ".js"   { $response.ContentType = "text/javascript; charset=utf-8" }
                    ".css"  { $response.ContentType = "text/css; charset=utf-8" }
                    ".png"  { $response.ContentType = "image/png" }
                    ".jpg"  { $response.ContentType = "image/jpeg" }
                    ".csv"  { $response.ContentType = "text/csv; charset=utf-8" }
                    default { $response.ContentType = "application/octet-stream" }
                }
                
                $response.ContentLength64 = $content.Length
                $response.OutputStream.Write($content, 0, $content.Length)
                
                Write-Host "✅ $($request.HttpMethod) $url - 200 OK ($($content.Length) bytes)" -ForegroundColor Green
            } catch {
                Write-Host "❌ ファイル読み込みエラー $url : $($_.Exception.Message)" -ForegroundColor Red
                $response.StatusCode = 500
            }
        } else {
            Write-Host "❌ $($request.HttpMethod) $url - 404 Not Found" -ForegroundColor Red
            $response.StatusCode = 404
            $errorContent = [System.Text.Encoding]::UTF8.GetBytes("404 - File Not Found")
            $response.OutputStream.Write($errorContent, 0, $errorContent.Length)
        }
        
        $response.Close()
    }
} catch {
    Write-Host "❌ サーバーエラー: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($listener.IsListening) {
        $listener.Stop()
    }
    Write-Host "🔧 サーバー停止" -ForegroundColor Yellow
}
