"""
DOCX Rebuilder GUI - Windows GUI with progress bar for document rebuilding.
"""

import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional, Callable


class ProgressWindow:
    """A simple progress window for the rebuild process."""

    def __init__(self, input_file: Optional[Path] = None):
        self.input_file = input_file
        self.root = tk.Tk()
        self.root.title("DOCX Rebuilder")
        self.root.geometry("500x280")
        self.root.resizable(False, False)

        # Center the window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 500) // 2
        y = (self.root.winfo_screenheight() - 280) // 2
        self.root.geometry(f"500x280+{x}+{y}")

        # Set icon if available
        try:
            self.root.iconbitmap(default='')
        except Exception:
            pass

        self._setup_ui()
        self._worker_thread: Optional[threading.Thread] = None
        self._cancelled = False

    def _setup_ui(self):
        """Set up the user interface."""
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="DOCX Rebuilder",
            font=('Segoe UI', 16, 'bold')
        )
        title_label.pack(pady=(0, 5))

        # Subtitle
        subtitle_label = ttk.Label(
            main_frame,
            text="Fix Word document formatting",
            font=('Segoe UI', 10),
            foreground='gray'
        )
        subtitle_label.pack(pady=(0, 15))

        # File selection frame
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(file_frame, text="Input file:").pack(anchor=tk.W)

        file_entry_frame = ttk.Frame(file_frame)
        file_entry_frame.pack(fill=tk.X, pady=(5, 0))

        self.file_var = tk.StringVar(value=str(self.input_file) if self.input_file else "")
        self.file_entry = ttk.Entry(file_entry_frame, textvariable=self.file_var, width=50)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        browse_btn = ttk.Button(file_entry_frame, text="Browse...", command=self._browse_file)
        browse_btn.pack(side=tk.RIGHT, padx=(10, 0))

        # Status label
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            font=('Segoe UI', 9)
        )
        self.status_label.pack(anchor=tk.W, pady=(0, 5))

        # Progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=460,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 15))

        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)

        self.rebuild_btn = ttk.Button(
            btn_frame,
            text="Rebuild Document",
            command=self._start_rebuild,
            width=20
        )
        self.rebuild_btn.pack(side=tk.LEFT)

        self.cancel_btn = ttk.Button(
            btn_frame,
            text="Cancel",
            command=self._cancel,
            state=tk.DISABLED,
            width=15
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=(10, 0))

        close_btn = ttk.Button(
            btn_frame,
            text="Close",
            command=self.root.destroy,
            width=15
        )
        close_btn.pack(side=tk.RIGHT)

        # Auto-start if file was provided
        if self.input_file and self.input_file.exists():
            self.root.after(500, self._start_rebuild)

    def _browse_file(self):
        """Open file browser dialog."""
        filepath = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[
                ("Word Documents", "*.docx"),
                ("All Files", "*.*")
            ]
        )
        if filepath:
            self.file_var.set(filepath)

    def _update_progress(self, value: float, status: str = ""):
        """Update progress bar and status (thread-safe)."""
        def update():
            self.progress_var.set(value)
            if status:
                self.status_var.set(status)
        self.root.after(0, update)

    def _start_rebuild(self):
        """Start the rebuild process in a background thread."""
        filepath = self.file_var.get().strip()
        if not filepath:
            messagebox.showerror("Error", "Please select a file first.")
            return

        input_path = Path(filepath)
        if not input_path.exists():
            messagebox.showerror("Error", f"File not found:\n{filepath}")
            return

        if not input_path.suffix.lower() == '.docx':
            messagebox.showerror("Error", "Please select a .docx file.")
            return

        # Disable controls during processing
        self.rebuild_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.file_entry.config(state=tk.DISABLED)
        self._cancelled = False

        # Start worker thread
        self._worker_thread = threading.Thread(
            target=self._rebuild_worker,
            args=(input_path,),
            daemon=True
        )
        self._worker_thread.start()

    def _rebuild_worker(self, input_path: Path):
        """Worker thread that performs the actual rebuild."""
        try:
            from .rebuilder import DocxRebuilder

            output_path = input_path.with_stem(input_path.stem + '_rebuilt')

            # Phase 1: Parsing
            self._update_progress(10, "Parsing document...")
            if self._cancelled:
                return

            rebuilder = DocxRebuilder(input_path, output_path)
            rebuilder._parse_document()

            if self._cancelled:
                return

            # Phase 2: Analyzing
            self._update_progress(30, "Analyzing formatting...")
            rebuilder._analyze_formatting()

            if self._cancelled:
                return

            # Phase 3: Initializing handlers
            self._update_progress(50, "Processing numbering and references...")
            rebuilder._initialize_handlers()

            if self._cancelled:
                return

            # Phase 4: Rebuilding
            self._update_progress(70, "Rebuilding document...")
            rebuilder._rebuild_document(
                normalize_numbering=True,
                normalize_styles=True,
                preserve_cross_refs=True
            )

            if self._cancelled:
                return

            # Done
            self._update_progress(100, "Complete!")

            # Show success message
            def show_success():
                messagebox.showinfo(
                    "Success",
                    f"Document rebuilt successfully!\n\nSaved to:\n{output_path}"
                )
            self.root.after(0, show_success)

        except Exception as e:
            def show_error():
                messagebox.showerror("Error", f"Rebuild failed:\n{str(e)}")
            self.root.after(0, show_error)
            self._update_progress(0, f"Error: {str(e)}")

        finally:
            # Re-enable controls
            def reset_controls():
                self.rebuild_btn.config(state=tk.NORMAL)
                self.cancel_btn.config(state=tk.DISABLED)
                self.file_entry.config(state=tk.NORMAL)
            self.root.after(0, reset_controls)

    def _cancel(self):
        """Cancel the current operation."""
        self._cancelled = True
        self.status_var.set("Cancelled")
        self.cancel_btn.config(state=tk.DISABLED)

    def run(self):
        """Run the GUI application."""
        self.root.mainloop()


def run_gui(input_file: Optional[str] = None):
    """Launch the GUI application."""
    path = Path(input_file) if input_file else None
    app = ProgressWindow(path)
    app.run()


if __name__ == '__main__':
    input_file = sys.argv[1] if len(sys.argv) > 1 else None
    run_gui(input_file)
