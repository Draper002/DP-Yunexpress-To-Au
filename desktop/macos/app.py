from __future__ import annotations

import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import messagebox


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from desktop.windows.app import App as SharedDesktopApp


class MacApp(SharedDesktopApp):
    """macOS adapter for the shared desktop fulfillment interface."""

    def __init__(self, root: tk.Tk):
        super().__init__(root)
        self.root.title("DP International Fulfillment To Au - macOS")

    @staticmethod
    def reveal_in_finder(path: Path) -> None:
        subprocess.Popen(["open", str(path)])

    def open_output_dir(self):
        if not self.refresh_output_dir():
            return
        path = Path(self.output_dir.get())
        path.mkdir(parents=True, exist_ok=True)
        self.reveal_in_finder(path)

    def open_last_result_dir(self):
        if not self.last_result_dir or not self.last_result_dir.exists():
            messagebox.showwarning("暂无结果", "当前还没有可打开的成功结果文件夹。")
            return
        self.reveal_in_finder(self.last_result_dir)


if __name__ == "__main__":
    app_root = tk.Tk()
    MacApp(app_root)
    app_root.mainloop()
