#!/usr/bin/env python3
"""
Anabim_Launcher.py – Lanceur graphique pour l’export IFC → Excel
================================================================

Fonction :
  • Demande à l’utilisateur de sélectionner un **dossier contenant des IFC**.
  • Exécute le script CLI `multi_anabim.py` sur ce dossier.
  • Affiche un message « Terminé » ou l’erreur capturée.

Empaquetage en exécutable :
  ```powershell
  python -m PyInstaller --onefile --noconsole \
      --add-data "multi_anabim.py;." \
      --collect-data ifcopenshell \
      Anabim_Launcher.py
  ```

 * `--add-data` colle le script CLI à la racine du bundle.
 * `--collect-data ifcopenshell` embarque les dossiers `express` (schémas .exp).
"""
from __future__ import annotations

import os
import runpy
import subprocess
import sys
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox

# --------------------------------------------------------------------------- #
# Configuration                                                               #
# --------------------------------------------------------------------------- #
SCRIPT_NAME = "multi_anabim.py"   # script CLI à déclencher

# Import “fantômes” pour que PyInstaller embarque les dépendances du script CLI
try:
    import pandas  # noqa: F401
    import openpyxl  # noqa: F401
    # ifcopenshell : on laisse PyInstaller le collecter via --collect-data
except ModuleNotFoundError:
    pass  # Lors du dev, ces libs sont déjà dans l'env ; on ignore sinon

# --------------------------------------------------------------------------- #
# Boîte de sélection de dossier                                               #
# --------------------------------------------------------------------------- #

def choose_folder() -> Optional[Path]:
    """Retourne le Path du dossier choisi ou None si annulation."""
    root = tk.Tk()
    root.withdraw()
    folder_str = filedialog.askdirectory(title="Sélectionner le dossier contenant les IFC")
    root.destroy()
    return Path(folder_str) if folder_str else None

# --------------------------------------------------------------------------- #
# Exécution du script CLI                                                     #
# --------------------------------------------------------------------------- #

def run_cli(folder: Path) -> None:
    script_path = Path(__file__).with_name(SCRIPT_NAME)
    if not script_path.exists():
        raise FileNotFoundError(f"Script principal introuvable : {script_path}")

    if getattr(sys, "frozen", False):  # EXE PyInstaller
        # Définit la variable pour les schémas EXPRESS avant d'importer ifcopenshell
        exp_dir = Path(sys._MEIPASS) / "ifcopenshell" / "express"  # type: ignore[attr-defined]
        os.environ.setdefault("IFCOPENSHELL_SCHEMA_PATH", str(exp_dir))
        # Exécution in-process du script CLI
        sys.argv = [script_path.name, str(folder)]
        runpy.run_path(str(script_path), run_name="__main__")
    else:  # exécution depuis Python/VS Code
        subprocess.run([sys.executable, str(script_path), str(folder)], check=True)

# --------------------------------------------------------------------------- #
# Point d’entrée                                                              #
# --------------------------------------------------------------------------- #

def main() -> None:
    folder = choose_folder()
    if folder is None:
        messagebox.showinfo("Annulé", "Aucun dossier sélectionné.")
        return

    try:
        run_cli(folder)
        messagebox.showinfo("Terminé", "Analyse IFC → Excel terminée !")
    except Exception as err:  # noqa: BLE001
        messagebox.showerror("Erreur", f"Le traitement a échoué.\n{err}")


if __name__ == "__main__":
    main()
