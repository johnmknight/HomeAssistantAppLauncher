import ctypes
import time
import sys

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
user32 = ctypes.WinDLL("user32")

VIRTUAL_KEYS = {
    **{f"VK_{chr(i)}": i for i in range(0x41, 0x5B)},
    "VK_0": 0x30, "VK_1": 0x31, "VK_2": 0x32, "VK_3": 0x33, "VK_4": 0x34,
    "VK_5": 0x35, "VK_6": 0x36, "VK_7": 0x37, "VK_8": 0x38, "VK_9": 0x39,
    **{f"VK_F{i}": 0x70 + i for i in range(1, 13)},
    "VK_BACK": 0x08, "VK_TAB": 0x09, "VK_RETURN": 0x0D, "VK_SHIFT": 0x10,
    "VK_CONTROL": 0x11, "VK_MENU": 0x12, "VK_PAUSE": 0x13, "VK_CAPITAL": 0x14,
    "VK_ESCAPE": 0x1B, "VK_SPACE": 0x20, "VK_PRIOR": 0x21, "VK_NEXT": 0x22,
    "VK_END": 0x23, "VK_HOME": 0x24, "VK_LEFT": 0x25, "VK_UP": 0x26,
    "VK_RIGHT": 0x27, "VK_DOWN": 0x28, "VK_SELECT": 0x29, "VK_PRINT": 0x2A,
    "VK_SNAPSHOT": 0x2C, "VK_INSERT": 0x2D, "VK_DELETE": 0x2E, "VK_HELP": 0x2F,
    "VK_VOLUME_MUTE": 0xAD, "VK_VOLUME_DOWN": 0xAE, "VK_VOLUME_UP": 0xAF,
    "VK_MEDIA_NEXT_TRACK": 0xB0, "VK_MEDIA_PREV_TRACK": 0xB1,
    "VK_MEDIA_STOP": 0xB2, "VK_MEDIA_PLAY_PAUSE": 0xB3,
    "VK_LAUNCH_MEDIA_SELECT": 0xB5
}

def simulate_key_press(key_name):
    print(f"Simulating key press: {key_name}")
    vk_code = VIRTUAL_KEYS.get(key_name)
    if vk_code is None:
        raise ValueError(f"Invalid virtual key name: {key_name}")
    user32.keybd_event(vk_code, 0, KEYEVENTF_EXTENDEDKEY, 0)
    time.sleep(0.05)
    user32.keybd_event(vk_code, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
    sys.exit()

if __name__ == "__main__":
    key = sys.argv[1] if len(sys.argv) > 1 else "VK_VOLUME_UP"
    simulate_key_press(key)
