import PyInstaller.__main__
import os
import platform

# Determine the operating system
current_os = platform.system()

os_folder_name = None
match current_os:
    case "Linux":
        os_folder_name = "linux"
    case "Windows":
        os_folder_name = "windows"
    case "Darwin":
        os_folder_name = "mac"
    case _:
        os_folder_name = "other"

# Create a directory path based on the operating system
dist_path = os.path.join(os.getcwd(), os_folder_name)

PyInstaller.__main__.run(
    [
        "app.py",
        "--onefile",
        "--windowed",
        "--clean",
        "--name",
        "Master Processor",
        "--noconfirm",
        "--distpath",
        dist_path,
    ]
)
