import tkinter as tk
from tkinter import ttk
import win32gui , win32ui , win32con , win32com.client , win32process , cv2 , numpy as np , os , psutil , sys
from PIL import Image, ImageTk

os.system('cls')


def list_windows():
    titles = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            titles.append((win32gui.GetWindowText(hwnd), hwnd))
    win32gui.EnumWindows(enum_handler, None)
    return titles


def capture_window(hwnd, resize=(200, 120)):
    try:
        left, top, right, bot = win32gui.GetClientRect(hwnd)
        width = right - left
        height = bot - top
        if width == 0 or height == 0:
            return None

        left, top = win32gui.ClientToScreen(hwnd, (left, top))
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)
        saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        img = np.frombuffer(bmpstr, dtype='uint8')
        img.shape = (bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        if resize:
            img = cv2.resize(img, resize)

        return img
    except:
        return None


PathFolder = []
def create_code_file(filename="main.bat"):
    CodeBat = "@echo off"
    for path in PathFolder:
        CodeBat += f'\nstart "" "{path}"'
        
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(CodeBat)

    create_shortcut(
    target_path=r"C:\\Users\\WINDOWS\\Desktop\\Techky\\main.bat",
    shortcut_name="open me",
    icon_path=r"C:\\Users\\WINDOWS\\Downloads\\Screenshot-2025-04-15-133810.ico"
    )


def get_exe_path_from_hwnd(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        proc = psutil.Process(pid)
        return proc.exe()
    except Exception as e:
        return f"Error: {e}"


def create_shortcut(target_path, shortcut_name, icon_path=None):
    desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    shortcut_path = os.path.join(desktop, f"{shortcut_name}.lnk")

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.WorkingDirectory = os.path.dirname(target_path)
    
    if icon_path:
        shortcut.IconLocation = icon_path
    
    shortcut.save()
    print(f"Shortcut created at: {shortcut_path}")

def restart_program():
    python = sys.executable
    os.execl(python, python, *sys.argv)


class WindowCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Techky")
        self.root.geometry("1000x700")

        self.canvas = tk.Canvas(root)
        self.v_scrollbar = tk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.h_scrollbar = tk.Scrollbar(root, orient="horizontal", command=self.canvas.xview)

        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        self.scrollable_frame = tk.Frame(self.canvas)

        #   SAVE BUTTON
        btnn = tk.Button(root, text="Save", command=create_code_file)
        btnn.pack(side=tk.TOP, padx=0)
        buttonrefresh = tk.Button(root, text="Refresh", command=restart_program)
        buttonrefresh.pack(side=tk.TOP, pady=3)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scrollbar.pack(side="right", fill="y")
        self.h_scrollbar.pack(side="bottom", fill="x")

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)           # Y axis
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shiftwheel)    # X axis

        self.selected_hwnd = None
        self.window_list = list_windows()
        self.build_preview_grid()


    def build_preview_grid(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        columns = 4
        for idx, (title, hwnd) in enumerate(self.window_list):
            frame = tk.Frame(self.scrollable_frame, bd=2, relief=tk.RIDGE)
            frame.grid(row=idx // columns, column=idx % columns, padx=5, pady=5)

            img = capture_window(hwnd)
            if img is None:
                continue

            pil_img = Image.fromarray(img)
            tk_img = ImageTk.PhotoImage(pil_img)

            label_title = tk.Label(frame, text=title[0:31], font=("Arial", 10), bg="#ccc", anchor="w")
            label_title.pack(fill=tk.X)

            label_img = tk.Label(frame, image=tk_img, bd=1, relief="solid", bg='white')
            label_img.image = tk_img
            label_img.pack()
            self.remove_highlight(label_img)

            label_img.bind("<Enter>", lambda e, lbl=label_img: self.highlight_hover(lbl))
            label_img.bind("<Leave>", lambda e, lbl=label_img: self.remove_highlight(lbl))
            label_img.bind("<Button-1>", lambda e, hwnd=hwnd: self.select_window(hwnd))
        

    def highlight_hover(self, label):
        label.config(highlightbackground='red', highlightthickness=2)


    def remove_highlight(self, label):
        label.config(highlightbackground='white', highlightthickness=2)


    def select_window(self, hwnd):
        self.selected_hwnd = hwnd
        title = win32gui.GetWindowText(hwnd)
        path = get_exe_path_from_hwnd(hwnd)
        print(f"🖥️  Window: {title}")
        print(f"📁 Path: {path}")
        PathFolder.append(str(path))
        print(f"[📁] Current folder: {PathFolder}")


    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


    def _on_shiftwheel(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")


# ===== Start GUI =====
if __name__ == "__main__":
    root = tk.Tk()
    app = WindowCaptureApp(root)
    root.mainloop()