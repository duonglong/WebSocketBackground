import sys

from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
        name = "HSS SOCKET",
        version = "1.1",
        description="Create websocket on port 8888",
        executables = [Executable("Socket_background.py", base = base)])