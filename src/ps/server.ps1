Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$SW_MINIMIZE = 6
$hwnd = [Win32]::GetConsoleWindow()
[Win32]::ShowWindow($hwnd, $SW_MINIMIZE) | Out-Null

$prefix = "http://localhost:8080/"
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
$listener.Start()

$profile = "$env:TEMP\edgeprofile"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$global:latestResult = ""
$global:targetImage = ""

function Start-Edge {
    if ($global:browser -and !$global:browser.HasExited) { return }

    Add-Type -AssemblyName System.Windows.Forms
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $width  = 600
    $height = 400
    $x = ($screen.Width  - $width) / 2
    $y = ($screen.Height - $height) / 2

    $profile = "$env:TEMP\edgeprofile"

    $global:browser = Start-Process "msedge.exe" `
        -ArgumentList "--user-data-dir=$profile",
                      "--new-window",
                      "--app=$prefix",
                      "--window-size=$width,$height",
                      "--window-position=$x,$y" `
        -PassThru
}

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $req  = $context.Request
    $resp = $context.Response

    $path = $req.Url.AbsolutePath.TrimStart("/")
    if ([string]::IsNullOrWhiteSpace($path)) { $path = "index.html" }

    # --- API: infer ---
    if ($path -eq "api/infer" -and $req.HttpMethod -eq "POST") {
        Start-Edge
        $json = '{"status":"accepted"}'
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $resp.ContentType = "application/json"
        $resp.OutputStream.Write($bytes,0,$bytes.Length)
        $resp.Close()
        continue
    }

    # --- API: result POST ---
    elseif ($path -eq "api/result" -and $req.HttpMethod -eq "POST") {
        $reader = New-Object IO.StreamReader($req.InputStream, $req.ContentEncoding)
        $body   = $reader.ReadToEnd()
        $global:latestResult = $body
        $bytes = [Text.Encoding]::UTF8.GetBytes($body)
        $reader.Close()
        $resp.ContentType = "application/json"
        $resp.OutputStream.Write($bytes,0,$bytes.Length)
        $resp.Close()
        try { Stop-Process -Id $browser.Id -Force } catch {}
        continue
    }

    # --- API: getresult ---
    elseif ($path -eq "api/getresult" -and $req.HttpMethod -eq "GET") {
        $bytes = [Text.Encoding]::UTF8.GetBytes($global:latestResult)
        $resp.ContentType = "application/json"
        $resp.OutputStream.Write($bytes,0,$bytes.Length)
        $resp.Close()
        $global:latestResult = ""
        continue
    }

    # --- Static file serving ---
    $file = Join-Path $root $path
    if (Test-Path $file -PathType Leaf) {
        $bytes = [IO.File]::ReadAllBytes($file)
        switch -Regex ($file) {
            '\.html$' { $resp.ContentType = "text/html; charset=utf-8" }
            '\.js$'   { $resp.ContentType = "application/javascript" }
            '\.css$'  { $resp.ContentType = "text/css" }
            '\.txt$'  { $resp.ContentType = "text/plain" }
            '\.wasm$' { $resp.ContentType = "application/wasm" }
            '\.onnx$' { $resp.ContentType = "application/octet-stream" }
            '\.bmp$'  { $resp.ContentType = "image/bmp" }
            '\.jpg$'  { $resp.ContentType = "image/jpeg" }
            '\.png$'  { $resp.ContentType = "image/png" }
            default   { $resp.ContentType = "application/octet-stream" }
        }
        $resp.ContentLength64 = $bytes.Length
        $resp.OutputStream.Write($bytes, 0, $bytes.Length)
        $resp.Close()
        continue
    }

    # --- Not found ---
    $resp.StatusCode = 404
    $resp.Close()
}
