"""
Usage Examples:

1. Using Python and command line:
   python url_f11.py https://spacedashboard.com/

2. Using curl to trigger via win_app_launcher:
   curl "http://192.168.4.47:8081/launch-url-to-f11?key=https://spacedashboard.com/"
"""

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
    result = subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        capture_output=True, text=True
    )

    if "OPENED_NEW" in result.stdout:
        print("Opened new window/tab. Waiting for browser to become active, then sending F11...")
        time.sleep(3)  # Wait for browser to launch (adjust as needed)
    else:
        print("Focused existing browser tab. Sending F11 for full screen...")
        time.sleep(1)  # Short delay for focus

    send_f11()
    print("Sent F11 to active window.")

if __name__ == "__main__":
    # Allow "key" param as well for web/launcher compatibility
    url = None
    if len(sys.argv) >= 2:
        url = sys.argv[1]
    else:
        # Check for key=... in environment or as fallback for launcher/curl compatibility
        import os
        url = os.environ.get("key", None)
        if not url:
            print("Usage: python url_f11.py <url>")
            sys.exit(1)
    open_or_fullscreen_url(url)
