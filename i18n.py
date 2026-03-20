"""
English / Hungarian UI strings for EML to PST Converter.
"""
from __future__ import annotations

import locale
import os

LANG_EN = "en"
LANG_HU = "hu"

# Combobox display names
LANGUAGE_NAMES = {
    LANG_EN: "English",
    LANG_HU: "Magyar",
}


def detect_default_lang() -> str:
    """Pick Hungarian if Windows/user locale is Hungarian."""
    try:
        loc = locale.getdefaultlocale()[0]
        if loc and str(loc).lower().startswith("hu"):
            return LANG_HU
    except Exception:
        pass
    env = os.environ.get("EML2PST_LANG", "").strip().lower()
    if env in ("hu", "hungarian", "magyar"):
        return LANG_HU
    if env in ("en", "english"):
        return LANG_EN
    return LANG_EN


def lang_from_display(name: str) -> str:
    if name == LANGUAGE_NAMES[LANG_HU]:
        return LANG_HU
    return LANG_EN


STRINGS: dict[str, dict[str, str]] = {
    LANG_EN: {
        "window_title": "EML to PST Converter",
        "folder_section": "Add Folder Having *.eml / *.emlx Files",
        "add_files": "Add Files",
        "file_pattern_label": "File Pattern:",
        "file_pattern_hint": "(Wildcards: * for any characters, ? for single character)",
        "save_pst_section": "Save in PST",
        "create_new_pst": "Create New PST File",
        "save_existing_pst": "Saving in Existing PST File",
        "remove_duplicates": "Remove Duplicate Content",
        "strict_date_preservation": (
            "Strict header dates (skip if Date/Received cannot be applied — optional)"
        ),
        # Explorer column: English "Date modified" vs Hungarian "Módosítás dátuma"
        "explorer_date_modified": "Date modified",
        "use_file_mtime": (
            "Match Explorer {explorer} in Outlook (native + fallback; on by default)"
        ),
        "eml_files_section": "EML/EMLX Files",
        "col_name": "EML/EMLX Name",
        "col_path": "Path",
        "col_size": "Size",
        "col_date": "Date Modified",
        "total_files": "Total Files: {n}",
        "destination_label": "Destination:",
        "browse_destination": "Browse Destination",
        "status_ready": "Ready",
        "btn_exit": "Exit",
        "btn_convert": "Convert",
        "language_label": "Language:",
        "context_remove": "Remove Selected",
        "context_clear_all": "Clear All",
        "dialog_select_folder": "Select Folder Containing EML/EMLX Files",
        "dialog_save_pst": "Save PST File",
        "dialog_open_pst": "Select Existing PST File",
        "dialog_filetype_pst": "Outlook PST Files",
        "dialog_all_files": "All Files",
        "confirm_exit_title": "Confirm",
        "confirm_exit_msg": "Conversion in progress. Exit anyway?",
        "warn_in_progress": "Conversion already in progress!",
        "warn_no_files": "No EML/EMLX files to convert!",
        "warn_no_destination": "Please select a destination path!",
        "err_dest_dir": "Destination directory does not exist: {path}",
        "err_scan": "Error scanning folder: {err}",
        "err_invalid_pattern": "Invalid file pattern. Only *.eml, *.emlx patterns are allowed.",
        "title_error": "Error",
        "title_warning": "Warning",
        "title_conversion_error": "Conversion Error",
        "msg_conversion_error": (
            "Error during conversion:\n{err}\n\n"
            "Make sure Microsoft Outlook is installed and working properly."
        ),
        "title_missing_dep": "Missing Dependency",
        "msg_missing_dep": (
            "The 'pywin32' library is required for PST conversion.\n\n"
            "Would you like to install it now?\n\n"
            "(This requires an internet connection)"
        ),
        "status_installing_pywin32": "Installing pywin32...",
        "title_install_ok": "Installation Complete",
        "msg_install_ok": (
            "pywin32 has been installed successfully!\n\n"
            "Please restart the application to use PST conversion."
        ),
        "title_install_fail": "Installation Failed",
        "msg_install_fail": (
            "Failed to install pywin32.\n\n"
            "Please run manually in command prompt:\npip install pywin32"
        ),
        "status_scanning": "Scanning folder...",
        "status_found": "Found {n} files",
        "status_connecting": "Connecting to Outlook...",
        "status_creating_pst": "Creating PST file...",
        "status_converting": "Converting: {name} ({cur}/{total})",
        "status_complete": "Conversion Complete",
        "title_complete": "Complete",
        "msg_complete": (
            "Conversion Complete!\n\n"
            "Converted: {converted} files\n"
            "Skipped: {skipped} files\n"
            "Errors: {errors} files"
        ),
        "note_pst_saved": "\n\nPST saved to: {path}",
        "note_errors": "\n\nErrors:\n",
        "note_skipped": "\n\nSkipped:\n",
        "note_more_errors": "\n... and {n} more errors",
        "note_more_skipped": "\n... and {n} more skipped",
    },
    LANG_HU: {
        "window_title": "EML – PST konverter",
        "folder_section": "Mappa hozzáadása *.eml / *.emlx fájlokkal",
        "add_files": "Fájlok hozzáadása",
        "file_pattern_label": "Fájlminta:",
        "file_pattern_hint": "(Helyettesítők: * több karakter, ? egy karakter)",
        "save_pst_section": "Mentés PST-be",
        "create_new_pst": "Új PST fájl létrehozása",
        "save_existing_pst": "Meglévő PST fájlba mentés",
        "remove_duplicates": "Duplikátumok eltávolítása",
        "strict_date_preservation": (
            "Szigorú fejléc dátumok (kihagyás, ha Date/Received nem alkalmazható — opcionális)"
        ),
        "explorer_date_modified": "Módosítás dátuma",
        "use_file_mtime": (
            "Outlook idő egyezzen a Tallózó {explorer} oszlopával (natív + tartalék; alapból be)"
        ),
        "eml_files_section": "EML/EMLX fájlok",
        "col_name": "EML/EMLX név",
        "col_path": "Útvonal",
        "col_size": "Méret",
        "col_date": "Módosítás dátuma",
        "total_files": "Összes fájl: {n}",
        "destination_label": "Cél:",
        "browse_destination": "Cél tallózása",
        "status_ready": "Kész",
        "btn_exit": "Kilépés",
        "btn_convert": "Konvertálás",
        "language_label": "Nyelv:",
        "context_remove": "Kijelöltek eltávolítása",
        "context_clear_all": "Összes törlése",
        "dialog_select_folder": "Mappa kiválasztása EML/EMLX fájlokkal",
        "dialog_save_pst": "PST fájl mentése",
        "dialog_open_pst": "Meglévő PST fájl kiválasztása",
        "dialog_filetype_pst": "Outlook PST fájlok",
        "dialog_all_files": "Minden fájl",
        "confirm_exit_title": "Megerősítés",
        "confirm_exit_msg": "Konvertálás folyamatban. Biztosan kilép?",
        "warn_in_progress": "A konvertálás már folyamatban van!",
        "warn_no_files": "Nincs konvertálandó EML/EMLX fájl!",
        "warn_no_destination": "Válasszon célútvonalat!",
        "err_dest_dir": "A célmappa nem létezik: {path}",
        "err_scan": "Hiba a mappa beolvasásakor: {err}",
        "err_invalid_pattern": "Érvénytelen minta. Csak *.eml és *.emlx minták engedélyezettek.",
        "title_error": "Hiba",
        "title_warning": "Figyelmeztetés",
        "title_conversion_error": "Konvertálási hiba",
        "msg_conversion_error": (
            "Hiba a konvertálás során:\n{err}\n\n"
            "Ellenőrizze, hogy a Microsoft Outlook telepítve és működőképes-e."
        ),
        "title_missing_dep": "Hiányzó összetevő",
        "msg_missing_dep": (
            "A PST konvertáláshoz a 'pywin32' könyvtár szükséges.\n\n"
            "Telepítse most?\n\n"
            "(Internetkapcsolat szükséges)"
        ),
        "status_installing_pywin32": "pywin32 telepítése...",
        "title_install_ok": "Telepítés kész",
        "msg_install_ok": (
            "A pywin32 sikeresen települt.\n\n"
            "Indítsa újra az alkalmazást a PST konvertáláshoz."
        ),
        "title_install_fail": "Telepítés sikertelen",
        "msg_install_fail": (
            "A pywin32 telepítése nem sikerült.\n\n"
            "Parancssorból futtassa: pip install pywin32"
        ),
        "status_scanning": "Mappa beolvasása...",
        "status_found": "{n} fájl találva",
        "status_connecting": "Kapcsolódás az Outlookhoz...",
        "status_creating_pst": "PST fájl létrehozása...",
        "status_converting": "Konvertálás: {name} ({cur}/{total})",
        "status_complete": "Konvertálás kész",
        "title_complete": "Kész",
        "msg_complete": (
            "Konvertálás kész!\n\n"
            "Konvertálva: {converted} fájl\n"
            "Kihagyva: {skipped} fájl\n"
            "Hibák: {errors} fájl"
        ),
        "note_pst_saved": "\n\nPST mentve ide: {path}",
        "note_errors": "\n\nHibák:\n",
        "note_skipped": "\n\nKihagyva:\n",
        "note_more_errors": "\n... és még {n} hiba",
        "note_more_skipped": "\n... és még {n} kihagyott",
    },
}


def t(lang: str, key: str, **kwargs) -> str:
    """Translate a string; falls back to English if key missing."""
    table = STRINGS.get(lang) or STRINGS[LANG_EN]
    s = table.get(key)
    if s is None:
        s = STRINGS[LANG_EN].get(key, key)
    if kwargs:
        return s.format(**kwargs)
    return s
