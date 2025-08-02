import sys
import subprocess
import time

def main():
    if len(sys.argv) < 2:
        print("Usage: python youtube_fullscreen_helper.py <youtube_url>")
        sys.exit(1)

    youtube_url = sys.argv[1]

    # Update these as needed for your setup:
    LAUNCHER_BASE = "http://localhost:8080"

    # 1. Call url-to-f11 endpoint to open/focus the video and send F11
    url_to_f11_endpoint = f'{LAUNCHER_BASE}/launch-url-to-f11?key={youtube_url}'
    print(f"Calling: {url_to_f11_endpoint}")
    subprocess.run(["curl", url_to_f11_endpoint])

    # Wait a bit to ensure the video/browser is ready for the next key
    time.sleep(3)

    # 2. Call vkey endpoint to send 'f' to the player (YouTube full screen)
    vkey_endpoint = f'{LAUNCHER_BASE}/launch-vkey?key=f'
    print(f"Calling: {vkey_endpoint}")
    subprocess.run(["curl", vkey_endpoint])

if __name__ == "__main__":
    main()
