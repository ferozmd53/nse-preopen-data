import subprocess

command = r'irm https://get.activated.win | iex'

subprocess.Popen(
    ["powershell.exe", "-NoExit", "-Command", command]
)
