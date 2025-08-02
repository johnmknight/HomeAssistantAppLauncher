import sys
import subprocess
import time
import ctypes

def send_f11():
    VK_F11 = 0x7A
    KEYEVENTF_KEYUP = 0x0002
    # Key down
    ctypes.windll.user32.keybd_event(VK_F11, 0, 0, 0)
    time.sleep(0.05)
    # Key up
    ctypes.windll.user32.keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0)

def open_or_fullscreen_url(url):
    url_lc = url.lower()

    # PowerShell script checks for open tabs and focuses or opens new window
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
if (-not $found) {{ Write-Output "OPENED_NEW" }} else {{ Write-Output "FOCUSED" }}
"""
    # Run the PowerShell script and capture output
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        capture_output=True, text=True
    )

    # If a new window was opened, wait and send F11
    if "OPENED_NEW" in result.stdout:
        print("Opened new window, sending F11 to fullscreen...")
        # Wait for the browser to launch and become active
        time.sleep(3)
        send_f11()
    else:
        print("Focused existing browser tab (no F11 sent).")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python url_f11.py <url>")
        sys.exit(1)
    url = sys.argv[1]
    open_or_fullscreen_url(url)
