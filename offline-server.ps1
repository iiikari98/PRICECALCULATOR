$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 5173

function Get-ContentType($path) {
  switch ([System.IO.Path]::GetExtension($path).ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8" }
    ".js" { "text/javascript; charset=utf-8" }
    ".css" { "text/css; charset=utf-8" }
    ".json" { "application/json; charset=utf-8" }
    ".png" { "image/png" }
    ".jpg" { "image/jpeg" }
    ".jpeg" { "image/jpeg" }
    ".svg" { "image/svg+xml" }
    ".ttf" { "font/ttf" }
    ".woff" { "font/woff" }
    ".woff2" { "font/woff2" }
    ".pdf" { "application/pdf" }
    default { "application/octet-stream" }
  }
}

while ($true) {
  try {
    $listener = [System.Net.HttpListener]::new()
    $prefix = "http://127.0.0.1:$port/"
    $listener.Prefixes.Add($prefix)
    $listener.Start()
    break
  } catch {
    if ($listener) {
      $listener.Close()
    }
    $port += 1
    if ($port -gt 5190) {
      throw "Could not start local server on ports 5173-5190."
    }
  }
}

$url = "http://127.0.0.1:$port/"
Write-Host "Quotation Calculator is running:"
Write-Host $url
Write-Host "Keep this window open while using the tool. Press Ctrl+C to stop."
Start-Process $url

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $requestPath = [System.Uri]::UnescapeDataString($context.Request.Url.AbsolutePath.TrimStart("/"))
    if ([string]::IsNullOrWhiteSpace($requestPath)) {
      $requestPath = "index.html"
    }

    $fullPath = Join-Path $root $requestPath
    $resolvedRoot = [System.IO.Path]::GetFullPath($root)
    $resolvedPath = [System.IO.Path]::GetFullPath($fullPath)

    if (-not $resolvedPath.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
      $context.Response.StatusCode = 403
      $context.Response.Close()
      continue
    }

    if (-not [System.IO.File]::Exists($resolvedPath)) {
      $resolvedPath = Join-Path $root "index.html"
    }

    $bytes = [System.IO.File]::ReadAllBytes($resolvedPath)
    $context.Response.ContentType = Get-ContentType $resolvedPath
    $context.Response.ContentLength64 = $bytes.Length
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $context.Response.OutputStream.Close()
  }
} finally {
  if ($listener) {
    $listener.Stop()
    $listener.Close()
  }
}
