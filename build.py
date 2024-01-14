import PyInstaller.__main__
import os
import platform
import subprocess

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

# create a path to the icon
mac_icon = "icon.icns"
win_icon = "icon.ico"
icon_name = mac_icon if current_os == "Darwin" else win_icon
icon_path = os.path.join("icons", icon_name)

# Specify the path to your Bash script
bash_script_path = "scripts/generate_icons.sh"
subprocess.run(["bash", bash_script_path])


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
        "--icon",
        icon_path,
    ]
)
