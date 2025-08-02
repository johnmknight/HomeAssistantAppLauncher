Set WshShell = CreateObject("WScript.Shell")
' Running the Python script silently without opening a command prompt window
WshShell.Run """C:\Users\john_\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\python.exe"" ""C:\Users\john_\HomeAssistantAppLauncher\win_app_launcher.py""", 0, False