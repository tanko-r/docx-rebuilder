"""
Windows Context Menu Installer for DOCX Rebuilder.

This script adds a "Rebuild with DOCX Rebuilder" option to the
right-click context menu for .docx files on Windows.

Run as Administrator to install/uninstall.
"""

import sys
import os
import winreg
import ctypes
from pathlib import Path


def is_admin():
    """Check if the script is running with admin privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def get_python_executable():
    """Get the path to the Python executable."""
    return sys.executable


def get_script_path():
    """Get the path to the GUI launcher script."""
    # Look for the package installation
    try:
        import docx_rebuilder
        package_dir = Path(docx_rebuilder.__file__).parent
        return package_dir / "gui.py"
    except ImportError:
        # Fall back to relative path
        script_dir = Path(__file__).parent.parent
        return script_dir / "docx_rebuilder" / "gui.py"


def install_context_menu():
    """Install the context menu entry for .docx files."""
    python_exe = get_python_executable()
    script_path = get_script_path()

    if not script_path.exists():
        print(f"Error: GUI script not found at {script_path}")
        return False

    # Create the command
    # Use pythonw.exe for GUI (no console window)
    pythonw_exe = python_exe.replace('python.exe', 'pythonw.exe')
    if not Path(pythonw_exe).exists():
        pythonw_exe = python_exe

    command = f'"{pythonw_exe}" "{script_path}" "%1"'

    try:
        # Registry path for .docx files
        # Method 1: Add to .docx shell
        key_path = r"SOFTWARE\Classes\.docx\shell\DocxRebuilder"

        # Create the main key
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValue(key, "", winreg.REG_SZ, "Rebuild with DOCX Rebuilder")
        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f"{python_exe},0")
        winreg.CloseKey(key)

        # Create the command key
        cmd_key_path = key_path + r"\command"
        cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, cmd_key_path)
        winreg.SetValue(cmd_key, "", winreg.REG_SZ, command)
        winreg.CloseKey(cmd_key)

        # Method 2: Also add to Word.Document.12 (for files opened with Word)
        word_key_path = r"SOFTWARE\Classes\Word.Document.12\shell\DocxRebuilder"

        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, word_key_path)
            winreg.SetValue(key, "", winreg.REG_SZ, "Rebuild with DOCX Rebuilder")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f"{python_exe},0")
            winreg.CloseKey(key)

            cmd_key_path = word_key_path + r"\command"
            cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, cmd_key_path)
            winreg.SetValue(cmd_key, "", winreg.REG_SZ, command)
            winreg.CloseKey(cmd_key)
        except Exception as e:
            print(f"Warning: Could not add to Word.Document.12: {e}")

        print("Context menu installed successfully!")
        print(f"Python: {pythonw_exe}")
        print(f"Script: {script_path}")
        print("\nRight-click any .docx file to see 'Rebuild with DOCX Rebuilder'")
        return True

    except PermissionError:
        print("Error: Permission denied. Try running as Administrator.")
        return False
    except Exception as e:
        print(f"Error installing context menu: {e}")
        return False


def uninstall_context_menu():
    """Remove the context menu entry."""
    keys_to_remove = [
        r"SOFTWARE\Classes\.docx\shell\DocxRebuilder",
        r"SOFTWARE\Classes\Word.Document.12\shell\DocxRebuilder",
    ]

    success = True
    for key_path in keys_to_remove:
        try:
            # First delete the command subkey
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path + r"\command")
            except FileNotFoundError:
                pass

            # Then delete the main key
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            except FileNotFoundError:
                pass

        except Exception as e:
            print(f"Warning: Could not remove {key_path}: {e}")
            success = False

    if success:
        print("Context menu uninstalled successfully!")
    return success


def main():
    """Main entry point."""
    print("=" * 50)
    print("DOCX Rebuilder - Windows Context Menu Installer")
    print("=" * 50)
    print()

    if len(sys.argv) > 1:
        if sys.argv[1] in ['--uninstall', '-u', 'uninstall']:
            uninstall_context_menu()
            return
        elif sys.argv[1] in ['--help', '-h', 'help']:
            print("Usage:")
            print("  python install_context_menu.py           Install context menu")
            print("  python install_context_menu.py --uninstall   Remove context menu")
            return

    install_context_menu()


if __name__ == '__main__':
    main()
