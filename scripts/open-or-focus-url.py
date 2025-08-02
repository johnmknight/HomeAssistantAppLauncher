import sys
import subprocess

def open_or_focus_url(url):
    # Lowercase for comparison (in PowerShell, we'll .ToLower())
    url_lc = url.lower()
    ps_script = f"""
$url = "{url_lc}"
$found = $false
$shell = New-Object -ComObject Shell.Application
$windows = $shell.Windows()
foreach ($window in $windows) {{
    try {{
        $tabUrl = $window.LocationURL
        if ($tabUrl -and $tabUrl.ToLower().TrimEnd('/') -eq $url.TrimEnd('/')) {{
            $window.Visible = $true
            $window.Focus()
            $found = $true
            break
        }}
    }} catch {{}}
}}
if (-not $found) {{
    Start-Process "{url}"
}}
"""
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script])

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python open_or_focus_url.py <url>")
        sys.exit(1)
    url = sys.argv[1]
    open_or_focus_url(url) 