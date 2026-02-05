"""
DOCX Rebuilder GUI - Windows GUI with progress bar for document rebuilding.
"""

import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import Optional, Callable
from datetime import datetime


class ManifestWindow:
    """A popup window showing the rebuild manifest/summary."""

    def __init__(self, parent, output_path: Path, report: dict, suggestions: list):
        self.window = tk.Toplevel(parent)
        self.window.title("Rebuild Complete - Manifest")
        self.window.geometry("550x500")
        self.window.resizable(True, True)

        # Center on parent
        self.window.transient(parent)
        self.window.grab_set()

        # Make it modal
        x = parent.winfo_x() + (parent.winfo_width() - 550) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 500) // 2
        self.window.geometry(f"550x500+{x}+{y}")

        self._setup_ui(output_path, report, suggestions)

    def _setup_ui(self, output_path: Path, report: dict, suggestions: list):
        """Set up the manifest UI."""
        main_frame = ttk.Frame(self.window, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header with checkmark
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            header_frame,
            text="✓ Document Rebuilt Successfully",
            font=('Segoe UI', 14, 'bold'),
            foreground='green'
        ).pack(side=tk.LEFT)

        # Output file
        ttk.Label(
            main_frame,
            text=f"Saved to: {output_path.name}",
            font=('Segoe UI', 9),
            foreground='gray'
        ).pack(anchor=tk.W)

        ttk.Label(
            main_frame,
            text=f"Location: {output_path.parent}",
            font=('Segoe UI', 9),
            foreground='gray'
        ).pack(anchor=tk.W, pady=(0, 10))

        # Scrolled text for manifest
        manifest_frame = ttk.LabelFrame(main_frame, text="Rebuild Manifest", padding="10")
        manifest_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.manifest_text = scrolledtext.ScrolledText(
            manifest_frame,
            wrap=tk.WORD,
            font=('Consolas', 9),
            height=18,
            state=tk.NORMAL
        )
        self.manifest_text.pack(fill=tk.BOTH, expand=True)

        # Build manifest content
        manifest = self._build_manifest(report, suggestions)
        self.manifest_text.insert(tk.END, manifest)
        self.manifest_text.config(state=tk.DISABLED)

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)

        ttk.Button(
            btn_frame,
            text="Open File Location",
            command=lambda: self._open_location(output_path)
        ).pack(side=tk.LEFT)

        ttk.Button(
            btn_frame,
            text="Copy Manifest",
            command=lambda: self._copy_manifest(manifest)
        ).pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(
            btn_frame,
            text="Close",
            command=self.window.destroy
        ).pack(side=tk.RIGHT)

    def _build_manifest(self, report: dict, suggestions: list) -> str:
        """Build the manifest text content."""
        lines = []
        lines.append("=" * 50)
        lines.append("DOCX REBUILDER - CONVERSION MANIFEST")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("=" * 50)
        lines.append("")

        # Document Statistics
        lines.append("DOCUMENT STATISTICS")
        lines.append("-" * 30)
        lines.append(f"  Paragraphs processed:    {report.get('paragraph_count', 0)}")
        lines.append(f"  Tables preserved:        {report.get('table_count', 0)}")
        lines.append(f"  Bookmarks found:         {report.get('bookmark_count', 0)}")
        lines.append(f"  Cross-references:        {report.get('cross_reference_count', 0)}")
        lines.append("")

        # Headers/Footers
        headers = report.get('headers_count', 0)
        footers = report.get('footers_count', 0)
        sections = report.get('sections_count', 0)
        page_fields = report.get('page_number_fields', 0)

        lines.append("HEADERS & FOOTERS")
        lines.append("-" * 30)
        lines.append(f"  Headers preserved:       {headers}")
        lines.append(f"  Footers preserved:       {footers}")
        lines.append(f"  Sections:                {sections}")
        lines.append(f"  Page number fields:      {page_fields}")
        lines.append("")

        # Formatting Analysis
        lines.append("FORMATTING ANALYSIS")
        lines.append("-" * 30)
        lines.append(f"  Base font detected:      {report.get('base_font', 'Unknown')}")
        lines.append(f"  Base font size:          {report.get('base_font_size', 0)}pt")
        lines.append(f"  Default alignment:       {report.get('default_alignment', 'left')}")
        lines.append(f"  Indent unit:             {report.get('indent_unit', 720)} twips")
        lines.append(f"  Heading styles found:    {report.get('heading_styles_detected', 0)}")
        lines.append(f"  Numbering schemes:       {report.get('numbering_schemes', 0)}")
        lines.append("")

        # What was done
        lines.append("ACTIONS PERFORMED")
        lines.append("-" * 30)
        lines.append("  [✓] Parsed document structure")
        lines.append("  [✓] Analyzed formatting conventions")
        lines.append("  [✓] Normalized numbering schemes")
        lines.append("  [✓] Preserved cross-references")
        lines.append("  [✓] Preserved headers and footers")
        lines.append("  [✓] Preserved page numbering")
        lines.append("  [✓] Rebuilt clean OOXML structure")
        lines.append("  [✓] Preserved table formatting")
        lines.append("  [✓] Maintained character formatting")
        lines.append("")

        # Issues detected (if any)
        if suggestions:
            lines.append("FORMATTING CORRECTIONS APPLIED")
            lines.append("-" * 30)
            # Group by type
            indent_fixes = [s for s in suggestions if s.get('type') == 'indent_correction']
            font_fixes = [s for s in suggestions if s.get('type') == 'font_correction']

            if indent_fixes:
                lines.append(f"  Indentation normalized:  {len(indent_fixes)} instances")
            if font_fixes:
                lines.append(f"  Font inconsistencies:    {len(font_fixes)} instances")

            # Show first few details
            lines.append("")
            lines.append("  Details (first 5):")
            for s in suggestions[:5]:
                lines.append(f"    - {s.get('message', 'Unknown correction')}")
            if len(suggestions) > 5:
                lines.append(f"    ... and {len(suggestions) - 5} more corrections")
            lines.append("")

        # Footer
        lines.append("=" * 50)
        lines.append("The rebuilt document should now have consistent")
        lines.append("formatting that works properly with MS Word.")
        lines.append("=" * 50)

        return "\n".join(lines)

    def _open_location(self, path: Path):
        """Open the file location in Explorer."""
        import subprocess
        import platform

        if platform.system() == 'Windows':
            subprocess.run(['explorer', '/select,', str(path)])
        elif platform.system() == 'Darwin':  # macOS
            subprocess.run(['open', '-R', str(path)])
        else:  # Linux
            subprocess.run(['xdg-open', str(path.parent)])

    def _copy_manifest(self, manifest: str):
        """Copy manifest to clipboard."""
        self.window.clipboard_clear()
        self.window.clipboard_append(manifest)
        messagebox.showinfo("Copied", "Manifest copied to clipboard!", parent=self.window)


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
            self._update_progress(45, "Processing numbering and references...")
            rebuilder._initialize_handlers()

            if self._cancelled:
                return

            # Report what was found
            report = rebuilder.get_analysis_report()
            headers_count = report.get('headers_count', 0)
            footers_count = report.get('footers_count', 0)
            if headers_count or footers_count:
                self._update_progress(55, f"Found {headers_count} headers, {footers_count} footers...")

            if self._cancelled:
                return

            # Get suggestions before rebuild
            suggestions = []
            if rebuilder.analyzer:
                suggestions = rebuilder.analyzer.suggest_corrections()

            # Phase 4: Rebuilding
            self._update_progress(70, "Rebuilding document (preserving headers/footers)...")
            rebuilder._rebuild_document(
                normalize_numbering=True,
                normalize_styles=True,
                preserve_cross_refs=True
            )

            if self._cancelled:
                return

            # Done - get final report
            self._update_progress(100, "Complete!")
            final_report = rebuilder.get_analysis_report()

            # Show manifest window
            def show_manifest():
                ManifestWindow(self.root, output_path, final_report, suggestions)
            self.root.after(0, show_manifest)

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
