from pathlib import Path
import sys

def resource_path(rel: str) -> str:
    """
    Liefert einen Pfad, der sowohl im Dev-Modus als auch in einer
    PyInstaller-OneFile-EXE funktioniert.
    rel: z. B. "resources/style.qss" oder "config/mapping.json"
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        # Projektroot = Ordner mit app.py
        base = Path(__file__).resolve().parents[1]
    return str((base / rel).resolve())
