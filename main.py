import os
import threading
from pathlib import Path
from tkinter import messagebox, filedialog
import customtkinter as ctk
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import ImageTk

# optional libs for better Arabic/Persian shaping
try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    SHAPING_AVAILABLE = True
except Exception:
    SHAPING_AVAILABLE = False

ENCODINGS_TO_TRY = ["utf-8", "utf-8-sig", "cp1256", "cp1252", "latin-1"]
APP_VERSION = "1.0.0"


class BatchMatlabToPdfApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Batch Folder -> PDF v{APP_VERSION}")
        self.iconpath = ImageTk.PhotoImage(file=self.resource_path(os.path.join("assets", "icon.png")))
        self.wm_iconbitmap()
        self.iconphoto(False, self.iconpath)
        self.update_idletasks()
        self.update_idletasks()
        width = 500
        height = 160
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.folder_var = ctk.StringVar()

        self._build_ui()

    def resource_path(self, relative_path):
        temp_dir = os.path.dirname(__file__)
        return os.path.join(temp_dir, relative_path)

    def _build_ui(self):
        frame = ctk.CTkFrame(self, corner_radius=8)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        label = ctk.CTkLabel(frame, text="Select folder with .m files:")
        label.grid(row=0, column=0, sticky="w", padx=(12, 6), pady=(8, 6))

        self.entry = ctk.CTkEntry(frame, textvariable=self.folder_var, width=420)
        self.entry.grid(row=1, column=0, sticky="w", padx=(12, 6))

        browse_btn = ctk.CTkButton(frame, text="Browse", width=100, command=self.browse_folder)
        browse_btn.grid(row=1, column=1, sticky="w", padx=(6, 12))

        convert_btn = ctk.CTkButton(
            frame, text="Convert All Codes to PDF", width=240, command=self.start_convert_thread
        )
        convert_btn.grid(row=2, column=0, columnspan=2, pady=(16, 6))

        # status
        self.status_label = ctk.CTkLabel(frame, text="")
        self.status_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=(12, 6), pady=(6, 0))

        # grid config
        frame.grid_columnconfigure(0, weight=1)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def start_convert_thread(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Please select a valid folder.")
            return
        # run in separate thread to keep UI responsive
        thread = threading.Thread(target=self.convert_all_to_pdf, args=(folder,), daemon=True)
        thread.start()

    def convert_all_to_pdf(self, folder_path):
        try:
            self._set_status("Scanning folder...")
            p = Path(folder_path)
            m_files = sorted([f for f in p.iterdir() if f.is_file() and f.suffix.lower() == ".m"])
            if not m_files:
                messagebox.showinfo("No files", "No .m files found in the selected folder.")
                self._set_status("")
                return

            output_pdf_path = p / "output.pdf"
            self._set_status(f"Creating PDF: {output_pdf_path.name}")

            # register a TTF font if available (try DejaVuSans)
            font_name = "Helvetica"
            registered_font = False
            try:
                # try typical locations or just rely on system installation path
                pdfmetrics.registerFont(
                    TTFont("DejaVuSans", self.resource_path(os.path.join("assets", "DejaVuSans.ttf")))
                )
                font_name = "DejaVuSans"
                registered_font = True
            except Exception:
                # fallback to built-in Helvetica
                font_name = "Helvetica"

            c = canvas.Canvas(str(output_pdf_path))
            page_width, page_height = c._pagesize

            left_margin = 40
            top_margin = page_height - 40
            line_height = 12
            font_size_title = 14  # filename
            font_size_code = 11  # code lines

            y = top_margin

            files_merged = 0

            for idx, file_path in enumerate(m_files):
                if idx > 0:  # skip new page for the very first file
                    c.showPage()
                    y = top_margin

                heading = file_path.name
                heading_to_write = self._maybe_shape_text(heading)

                c.setFont(font_name, font_size_title)
                c.drawString(left_margin, y, heading_to_write)
                y -= line_height * 2  # add one extra empty line before code

                c.setFont(font_name, font_size_code)

                # read file content with fallback encodings
                content = self._read_file_with_encodings(file_path)
                if content is None:
                    # write a note about failed read
                    failed_msg = f"[Could not read file: {file_path.name}]"
                    failed_msg = self._maybe_shape_text(failed_msg)
                    if y - line_height < 40:
                        c.showPage()
                        y = top_margin
                    c.drawString(left_margin, y, failed_msg)
                    y -= line_height
                    files_merged += 1
                    continue

                # write content line by line
                for raw_line in content.splitlines():
                    line = raw_line.rstrip()
                    if not line:
                        # write empty line spacing
                        y -= line_height * 0.6
                    else:
                        # optionally shape
                        line_to_write = self._maybe_shape_text(line)
                        # if line too long, wrap manually by characters to fit page width
                        max_chars = self._estimate_max_chars(font_size_code, page_width - 2 * left_margin)
                        wrapped_lines = self._wrap_text(line_to_write, max_chars)
                        for wline in wrapped_lines:
                            if y - line_height < 40:
                                c.showPage()
                                y = top_margin
                                c.setFont(font_name, font_size_code)
                            c.drawString(left_margin, y, wline)
                            y -= line_height
                # spacer between files
                y -= line_height * 0.8
                files_merged += 1

            c.save()
            self._set_status("Done")
            messagebox.showinfo(
                "Done", f"Merged {files_merged} files into '{output_pdf_path.name}' in the selected folder."
            )
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            self._set_status("")

    def _set_status(self, text):
        # update status label in main thread
        def updater():
            self.status_label.configure(text=text)

        try:
            self.after(1, updater)
        except Exception:
            pass

    def _read_file_with_encodings(self, path: Path):
        for enc in ENCODINGS_TO_TRY:
            try:
                with open(path, "r", encoding=enc) as f:
                    return f.read()
            except Exception:
                continue
        # final fallback: try binary decode with errors replaced
        try:
            with open(path, "rb") as f:
                data = f.read()
                return data.decode("utf-8", errors="replace")
        except Exception:
            return None

    def _maybe_shape_text(self, text: str) -> str:
        if SHAPING_AVAILABLE:
            try:
                reshaped = arabic_reshaper.reshape(text)
                bidi_text = get_display(reshaped)
                return bidi_text
            except Exception:
                return text
        else:
            return text

    def _estimate_max_chars(self, font_size: int, avail_width: float) -> int:
        # rough estimate: average char width approx 0.55 * font_size
        avg_char_width = 0.55 * font_size
        return max(20, int(avail_width / avg_char_width))

    def _wrap_text(self, text: str, max_chars: int):
        # naive wrapper that preserves words when possible
        if len(text) <= max_chars:
            return [text]
        parts = []
        current = ""
        for token in text.split(" "):
            if not current:
                current = token
            elif len(current) + 1 + len(token) <= max_chars:
                current += " " + token
            else:
                parts.append(current)
                current = token
        if current:
            parts.append(current)
        # further split extremely long parts
        finished = []
        for p in parts:
            if len(p) <= max_chars:
                finished.append(p)
            else:
                # split by fixed chunk
                for i in range(0, len(p), max_chars):
                    finished.append(p[i : i + max_chars])
        return finished


if __name__ == "__main__":
    app = BatchMatlabToPdfApp()
    app.mainloop()
