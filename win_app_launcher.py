import sys
import subprocess

# --- Auto-install pywin32 if missing ---
try:
    import win32com
except ImportError:
    print("Required module 'pywin32' not found. Attempting to install...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
        print("pywin32 installed. Please restart this script.")
        sys.exit(0)
    except Exception as e:
        print(f"Automatic install failed: {e}")
        print("Please run 'pip install pywin32' manually and then rerun this script.")
        sys.exit(1)

from win32com.client import Dispatch
import http.server
import socketserver
import socket
import os
import threading
import json
import time
import getpass
import shutil
import re
from urllib.parse import urlparse, parse_qs

# --- Configuration Constants ---
PERMANENT_SCRIPT_DIR = os.path.expanduser("~/HomeAssistantAppLauncher/")
CONFIG_FILE = os.path.join(PERMANENT_SCRIPT_DIR, "win_app_launcher_config.json")
STDOUT_LOG_PATH = os.path.join(PERMANENT_SCRIPT_DIR, "app_launcher_stdout.log")
STDERR_LOG_PATH = os.path.join(PERMANENT_SCRIPT_DIR, "app_launcher_stderr.log")
STARTUP_FOLDER = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
AUTOSTART_SCRIPT_NAME = "StartHomeAssistantAppLauncher.vbs"
AUTOSTART_SCRIPT_PATH = os.path.join(PERMANENT_SCRIPT_DIR, AUTOSTART_SCRIPT_NAME)
AUTOSTART_SHORTCUT_NAME = "HomeAssistantAppLauncher.lnk"

HANDLER_PORT = None
SCRIPT_CURRENT_PATH = os.path.abspath(__file__)
SCRIPT_PERMANENT_PATH = os.path.join(PERMANENT_SCRIPT_DIR, os.path.basename(__file__))

# --- Browser Handling with PowerShell ---
def activate_or_open_url(url, browser="default"):
    parsed = urlparse(url)
    domain = parsed.netloc.lower()
    if not domain:
        domain = url.lower()
    normalized_url = url.lower().rstrip('/')

    browser_paths = {
        "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "firefox": r"C:\Program Files\Mozilla Firefox\firefox.exe"
    }

    browser_exe = browser_paths.get(browser.lower(), None)
    browser_name = browser.lower() if browser.lower() in browser_paths else "default"

    ps_script = f"""
    $url = "{normalized_url}"
    $domain = "{domain}"
    $browser = "{browser_name}"

    function Get-Domain ($tabUrl) {{
        try {{
            $uri = [System.Uri]$tabUrl
            return $uri.Host.ToLower()
        }} catch {{
            return $tabUrl.ToLower()
        }}
    }}

    $shell = New-Object -ComObject Shell.Application
    $windows = $shell.Windows()
    $found = $false

    foreach ($window in $windows) {{
        if ($browser -eq "default" -or $window.FullName -like "*{browser_name}*") {{
            try {{
                $tabs = $window.Document.getElementsByTagName("HTML")
                foreach ($tab in $tabs) {{
                    $tabUrl = $window.LocationURL
                    $normalizedTabUrl = $tabUrl.ToLower().TrimEnd('/')
                    if ($normalizedTabUrl -eq $url) {{
                        $window.Navigate($url)
                        $window.Visible = $true
                        $found = $true
                        break
                    }}
                }}
            }} catch {{}}
            if ($found) {{ break }}
        }}
    }}

    if (-not $found) {{
        foreach ($window in $windows) {{
            if ($browser -eq "default" -or $window.FullName -like "*{browser_name}*") {{
                try {{
                    $tabs = $window.Document.getElementsByTagName("HTML")
                    foreach ($tab in $tabs) {{
                        $tabUrl = $window.LocationURL
                        $tabDomain = Get-Domain $tabUrl
                        if ($tabDomain -eq $domain) {{
                            $window.Navigate($tabUrl)
                            $window.Visible = $true
                            $found = $true
                            break
                        }}
                    }}
                }} catch {{}}
                if ($found) {{ break }}
            }}
        }}
    }}

    if (-not $found) {{
        if ($browser -eq "default") {{
            Start-Process $url
        }} else {{
            Start-Process -FilePath "{browser_exe}" -ArgumentList $url
        }}
    }}
    """

    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=True
        )
        with open(STDOUT_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] PowerShell output for {url}:\n{result.stdout}\n")
    except subprocess.CalledProcessError as e:
        error_msg = f"PowerShell error: {e.stderr}"
        with open(STDERR_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {error_msg}\n")
        raise Exception(error_msg)

# --- HTTP Request Handler ---
class AppHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        # Parse the path and query params
        url_parts = urlparse(self.path)
        path = url_parts.path.strip('/')
        query_params = parse_qs(url_parts.query)
        current_config = load_config()
        if path in current_config.get("apps", {}):
            app_info = current_config["apps"][path]
            item_type = app_info.get("type", "app")
            app_identifier = app_info["app_path"]

            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Received request for: {app_identifier} (type: {item_type})")
            try:
                if item_type == "app":
                    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Attempting to launch executable: {app_identifier}")
                    os.startfile(app_identifier)
                    success_msg = f"Launched: {app_identifier}"
                    if app_info.get("play_music", False):
                        time.sleep(2)
                        if "spotify" in app_identifier.lower():
                            ps_script = """
                            $spotify = New-Object -ComObject Spotify.Application
                            if ($spotify) { $spotify.Play() }
                            """
                            subprocess.run(
                                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                                capture_output=True,
                                text=True
                            )
                            print("Attempted to play music in Spotify.")
                        elif "wmplayer" in app_identifier.lower():
                            ps_script = """
                            $wmp = New-Object -ComObject WMPlayer.OCX
                            if ($wmp) { $wmp.controls.play() }
                            """
                            subprocess.run(
                                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                                capture_output=True,
                                text=True
                            )
                            print("Attempted to play music in Windows Media Player.")
                elif item_type == "uri":
                    uri_scheme = app_identifier[4:]
                    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Attempting to launch URI: {uri_scheme}")
                    subprocess.run(["start", "", uri_scheme], shell=True, check=True)
                    success_msg = f"Launched URI: {uri_scheme}"
                elif item_type == "website":
                    browser = app_info.get("browser", "default")
                    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Attempting to launch URL: {app_identifier} (browser: {browser})")
                    activate_or_open_url(app_identifier, browser)
                    success_msg = f"Opened URL: {app_identifier}"
                elif item_type == "python":
                    SCRIPTS_DIR = os.path.join(PERMANENT_SCRIPT_DIR, "scripts")
                    script_filename = app_identifier
                    script_path = os.path.join(SCRIPTS_DIR, script_filename)
                    if not os.path.abspath(script_path).startswith(os.path.abspath(SCRIPTS_DIR) + os.sep):
                        raise PermissionError("Python scripts must be inside the scripts directory.")
                    if not os.path.exists(script_path):
                        raise FileNotFoundError(f"Python script '{script_filename}' not found in scripts directory.")

                    # Check for 'key' parameter
                    key_param = query_params.get("key", [None])[0]
                    launch_args = [sys.executable, script_filename]
                    if key_param:
                        launch_args.append(key_param)
                        print(f"Passing argument to script: {key_param}")

                    print(f"Attempting to launch Python script via: {' '.join(launch_args)} (cwd={SCRIPTS_DIR})")
                    subprocess.Popen(launch_args, cwd=SCRIPTS_DIR)
                    success_msg = f"Launched Python script: {script_filename} (args: {key_param if key_param else 'none'})"

                self.send_response(200)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(f"<html><body>{success_msg}</body></html>".encode())
                print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {success_msg}")
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(f"<html><body>Error launching {app_identifier}: {e}</body></html>".encode())
                print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Error launching {app_identifier}: {e}")
        else:
            self.send_response(404)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(f"<html><body>Unknown command: {path}</body></html>".encode())
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Unknown command received: {path}")

# --- Helper Functions ---
def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

def find_available_port(start_port):
    for port in range(start_port, start_port + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("", port))
                return port
            except OSError:
                continue
    return None

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                print(f"Warning: Could not decode {CONFIG_FILE}. Starting with empty configuration.")
                return {"apps": {}, "port": None, "auto_start_enabled": False}
    return {"apps": {}, "port": None, "auto_start_enabled": False}

def save_config(config):
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

def setup_auto_run_windows(config):
    try:
        os.makedirs(PERMANENT_SCRIPT_DIR, exist_ok=True)
        os.makedirs(STARTUP_FOLDER, exist_ok=True)

        vbs_script_content = (
            "Set WshShell = CreateObject(\"WScript.Shell\")\n"
            "' Running the Python script silently without opening a command prompt window\n"
            "WshShell.Run \"\"\"{python_exe_path}\"\" \"\"{script_path}\"\"\", 0, False"
        )

        final_vbs_content = vbs_script_content.format(
            python_exe_path=os.path.normpath(sys.executable),
            script_path=os.path.normpath(SCRIPT_PERMANENT_PATH)
        )

        with open(AUTOSTART_SCRIPT_PATH, 'w') as f:
            f.write(final_vbs_content)

        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(os.path.join(STARTUP_FOLDER, AUTOSTART_SHORTCUT_NAME))
        shortcut.Targetpath = AUTOSTART_SCRIPT_PATH
        shortcut.WindowStyle = 7  # Minimized window style
        shortcut.save()

        print(f"\nSUCCESS: Auto-start VBScript created at: {AUTOSTART_SCRIPT_PATH}")
        print(f"SUCCESS: Auto-start shortcut created at: {os.path.join(STARTUP_FOLDER, AUTOSTART_SHORTCUT_NAME)}")

        config["auto_start_enabled"] = True
        save_config(config)
        return True
    except Exception as e:
        print(f"\nERROR: Problem setting up auto-run: {e}")
        return False

def disable_auto_run_windows(config):
    try:
        if os.path.exists(AUTOSTART_SCRIPT_PATH):
            os.remove(AUTOSTART_SCRIPT_PATH)
            print(f"Removed auto-start VBScript: {AUTOSTART_SCRIPT_PATH}")

        shortcut_path = os.path.join(STARTUP_FOLDER, AUTOSTART_SHORTCUT_NAME)
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            print(f"Removed auto-start shortcut: {shortcut_path}")

        config["auto_start_enabled"] = False
        save_config(config)
        return True
    except Exception as e:
        print(f"\nERROR: Problem removing auto-run: {e}")
        return False

def start_server(port):
    try:
        with socketserver.TCPServer(("", port), AppHandler) as httpd:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Serving at port {port}")
            httpd.serve_forever()
    except OSError as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ERROR: Could not start server on port {port}. Is it already in use? ({e})")
    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] An unexpected error occurred in server: {e}")

def main():
    global HANDLER_PORT
    os.makedirs(PERMANENT_SCRIPT_DIR, exist_ok=True)
    os.chdir(PERMANENT_SCRIPT_DIR)
    config = load_config()
    is_interactive = sys.stdin.isatty() and sys.stdout.isatty()

    if not is_interactive:
        sys.stdout = open(STDOUT_LOG_PATH, 'a', encoding='utf-8', buffering=1)
        sys.stderr = open(STDERR_LOG_PATH, 'a', encoding='utf-8', buffering=1)
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Home Assistant App Launcher started in background mode.")

    if is_interactive:
        print("="*70)
        print(">>> Home Assistant Windows App Launcher - Configuration Wizard <<<")
        print(f"Script location: {SCRIPT_PERMANENT_PATH}")
        print(f"Configuration and logs will be stored in: {PERMANENT_SCRIPT_DIR}")
        print("="*70)
        print("\nINSTRUCTIONS:")
        print(" - To add a Windows app:   Enter the full path to its .exe or shortcut (e.g. C:\\Program Files\\App\\app.exe)")
        print(" - To add a URI:           Enter a Windows URI (e.g. uri:ms-settings:sound)")
        print(" - To add a website:       Enter a full URL (e.g. https://www.google.com)")
        print(" - To add a Python script: Type 'python' (no quotes), then choose a script from the 'scripts/' folder")
        print("      (All Python scripts must be placed in the 'scripts/' folder inside your launcher directory.)")
        print(" - To enable/disable auto-start on login: Type 'auto-run'")
        print(" - To finish configuration: Type 'done'")
        print()

        local_ip = get_local_ip()
        print(f"Your Windows IP Address: {local_ip}")

        if config["port"]:
            HANDLER_PORT = config["port"]
            print(f"Using previously configured port: {HANDLER_PORT}")
        else:
            HANDLER_PORT = find_available_port(8080)
            if HANDLER_PORT:
                config["port"] = HANDLER_PORT
                save_config(config)
                print(f"Found and configured available port: {HANDLER_PORT}")
            else:
                print("ERROR: Could not find an available port. Please check your network or firewall settings.")
                input("Press Enter to exit...")
                return

        server_thread = threading.Thread(target=start_server, args=(HANDLER_PORT,))
        server_thread.daemon = True
        server_thread.start()

        print("\n--- Configure Apps, URIs, Websites, Python Scripts, and Auto-Start ---")
        while True:
            action_prompt = (
                "Enter an app's full path (.exe), URI (uri:...), website URL (https://...),\n"
                "or type 'python' to add a script from 'scripts/', 'auto-run' to configure autostart, or 'done' to finish: "
            )
            user_input = input(action_prompt).strip()

            if user_input.lower() == 'done':
                break
            elif user_input.lower() == 'auto-run':
                if config.get("auto_start_enabled"):
                    print("\nAuto-start is currently ENABLED.")
                    choice = input("Do you want to DISABLE auto-start? (yes/no): ").lower().strip()
                    if choice == 'yes':
                        disable_auto_run_windows(config)
                        print("\nIMPORTANT: Auto-start will be disabled on next reboot.")
                else:
                    print("\nAuto-start is currently DISABLED.")
                    print("This will create a VBScript and a shortcut in your Windows Startup folder.")
                    choice = input("Do you want to ENABLE auto-start? (yes/no): ").lower().strip()
                    if choice == 'yes':
                        if setup_auto_run_windows(config):
                            print("\nIMPORTANT: Restart your PC for auto-start to take full effect.")
                continue

            app_identifier_input = user_input
            item_type = "app"
            play_music = False
            browser = "default"

            # --- Begin Python script wizard support ---
            if app_identifier_input.lower() == "python":
                SCRIPTS_DIR = os.path.join(PERMANENT_SCRIPT_DIR, "scripts")
                os.makedirs(SCRIPTS_DIR, exist_ok=True)
                scripts = [f for f in os.listdir(SCRIPTS_DIR) if f.endswith('.py')]
                if not scripts:
                    print(f"No .py scripts found in '{SCRIPTS_DIR}'. Place your scripts there first!")
                    continue
                print("Available scripts in 'scripts/' directory:")
                for idx, script in enumerate(scripts, 1):
                    print(f"  {idx}. {script}")
                script_choice = input(f"Enter the script filename to add (e.g. {scripts[0]}): ").strip()
                if script_choice not in scripts:
                    print(f"Script '{script_choice}' not found in 'scripts/' directory.")
                    continue
                item_type = "python"
                app_identifier_input = script_choice
                base_name = os.path.splitext(script_choice)[0].lower().replace(" ", "-")
            elif app_identifier_input.lower().startswith("uri:"):
                if len(app_identifier_input) < 5 or ":" not in app_identifier_input[4:]:
                    print("Error: URI format invalid. Please use 'uri:scheme:' (e.g., uri:spotify:)")
                    continue
                item_type = "uri"
                base_name = app_identifier_input[4:].strip(':').lower().replace(" ", "-")
            elif app_identifier_input.lower().startswith("http://") or app_identifier_input.lower().startswith("https://"):
                try:
                    parsed_url = urlparse(app_identifier_input)
                    if not all([parsed_url.scheme, parsed_url.netloc]):
                        raise ValueError("Invalid URL format.")
                except ValueError:
                    print("Error: Invalid URL format. Please use 'http://' or 'https://' (e.g., https://www.google.com)")
                    continue
                item_type = "website"
                base_name = parsed_url.netloc.lower().replace("www.", "").replace(".", "-").split('/')[0]
                browser_choice = input("Which browser for this site? (chrome/edge/firefox/default) [default: default]: ").strip().lower()
                browser = browser_choice if browser_choice in ("chrome", "edge", "firefox", "default") else "default"
            else:
                if not os.path.exists(app_identifier_input):
                    print(f"Error: The path '{app_identifier_input}' does not appear to exist. Please enter a valid full path to an executable.")
                    continue
                if app_identifier_input.lower().endswith(".py"):
                    print("To add Python scripts, type 'python' at the prompt and select from scripts in the 'scripts/' directory.")
                    continue
                base_name = os.path.basename(app_identifier_input).split('.')[0].lower().replace(" ", "-")
                if "spotify" in base_name.lower() or "wmplayer" in base_name.lower():
                    play_music_choice = input(f"Do you want a separate URL to play music automatically when {app_identifier_input} is launched? (yes/no): ").lower().strip()
                    if play_music_choice == 'yes':
                        play_music = True

            app_url_name = f"launch-{base_name}"

            if app_url_name in config["apps"]:
                overwrite_choice = input(f"'{app_identifier_input}' already configured. Overwrite? (yes/no): ").lower().strip()
                if overwrite_choice != 'yes':
                    continue

            config["apps"][app_url_name] = {
                "type": item_type,
                "app_path": app_identifier_input,
                "play_music": play_music,
                "browser": browser if item_type == "website" else None
            }
            print(f"  Added '{app_identifier_input}' with URL path: /{app_url_name}")
            save_config(config)

        print("\n--- Home Assistant Configuration ---")
        print("Add the following to your Home Assistant 'configuration.yaml' file:")
        print("========================================================")
        if not config["apps"]:
            print("  No apps, URIs, or websites configured yet. Add them to generate Home Assistant commands.")
        else:
            print("rest_command:")
            for path, app_info in config["apps"].items():
                entity_id = f"{path.replace('-', '_')}_win"
                print(f"  {entity_id}:")
                print(f"    url: \"http://{local_ip}:{HANDLER_PORT}/{path}\"")
                print(f"    method: GET")
        print("========================================================")

        print(f"\nHTTP server is running at http://{local_ip}:{HANDLER_PORT}. Press Ctrl+C to stop.")
        print("This command prompt window needs to stay open for the server to run (unless auto-start is enabled).")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nServer stopped.")
            print("If auto-start is enabled and shortcut created, it's configured to resume on next login.")
            print("If auto-start is NOT enabled, you'll need to run 'python.exe win_app_launcher.py' again.")

    else:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Running in background service mode.")
        if not config.get("port"):
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ERROR: No port configured in {CONFIG_FILE}. Exiting background service.")
            return
        HANDLER_PORT = config["port"]
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Attempting to serve on port {HANDLER_PORT} from background.")
        start_server(HANDLER_PORT)

if __name__ == "__main__":
    main()
