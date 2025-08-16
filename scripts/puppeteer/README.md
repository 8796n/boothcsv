# Puppeteer screenshot helper for boothcsv

This small tool uses puppeteer-core to open the local HTTP server (http://localhost:8080/boothcsv.html) and take the same set of screenshots as `capture-all.ps1`.

Usage:

1. Install Node.js (>=16) and a Chrome/Chromium binary.
2. From `d:\dev\boothcsv\scripts\puppeteer` run:

```powershell
npm install
# optionally pass --chrome "C:\Program Files\Google\Chrome\Application\chrome.exe"
node screenshot.js --out "..\..\docs\images" --html "..\..\boothcsv.html" --chrome "C:\Path\to\chrome.exe"
```

Notes:
- If you don't provide --chrome, puppeteer-core will try to launch the system Chrome via default discovery.
- The script uses clip on elements when fitToSelector is provided; otherwise it captures the viewport.
- This is intended as a simpler, more reliable capture path if Selenium proves flaky.
