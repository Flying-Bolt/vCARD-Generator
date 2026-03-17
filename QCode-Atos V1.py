"""
vCard QR-Code Generator
Professioneller Dialog mit JSON-Speicherung, Festnetznummer und robuster Zwischenablage
"""

import io
import json
import logging
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Optional, Callable

# Logging konfigurieren
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

try:
    import qrcode
    from PIL import Image, ImageTk
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False
    logging.warning("Bibliotheken 'qrcode' oder 'Pillow' fehlen.")


# ─────────────────────────────────────────────
# Konstanten
# ─────────────────────────────────────────────
APP_TITLE   = "vCard QR-Code Generator"
my_BLUE   = "#001240"
WHITE       = "#FFFFFF"
ACCENT      = "#00A8E0"
BG_LIGHT    = "#F4F6F9"
BORDER      = "#D0D7DE"
TEXT_COLOR  = "#1A1A2E"
SUCCESS     = "#28A745"
ERROR_COLOR = "#DC3545"

FELDER = [
    ("vorname",      "Vorname",          ""),
    ("nachname",     "Nachname",         ""),
    ("position",     "Position / Titel", ""),
    ("mobiltelefon", "Mobiltelefon",     ""),
    ("festnetz",     "Festnetz",         ""),
    ("email",        "E-Mail",           ""),
    ("strasse",      "Straße",           ""),
    ("hausnummer",   "Hausnummer",       ""),
    ("plz",          "PLZ",              ""),
    ("ort",          "Ort",              ""),
    ("land",         "Land",             ""),
]

BASE_DIR = Path(__file__).parent
PROFILES_FILE = BASE_DIR / "QR-Code.json"


# ─────────────────────────────────────────────
# Daten-Management
# ─────────────────────────────────────────────
def lade_profile() -> Dict[str, Dict[str, str]]:
    if PROFILES_FILE.exists():
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            logging.error(f"Fehler beim Laden der Profile: {e}")
            return {}
    return {}


def speichere_profile(profiles: Dict[str, Dict[str, str]]) -> None:
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(profiles, f, ensure_ascii=False, indent=2)
    except IOError as e:
        logging.error(f"Fehler beim Speichern der Profile: {e}")


# ─────────────────────────────────────────────
# QR-Code Logik
# ─────────────────────────────────────────────
def erstelle_vcard(d: Dict[str, str]) -> str:
    lines = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"N:{d.get('nachname', '')};{d.get('vorname', '')};;;",
        f"FN:{d.get('vorname', '')} {d.get('nachname', '')}",
        "ORG:",
        f"TITLE:{d.get('position', '')}",
        f"TEL;TYPE=CELL:{d.get('mobiltelefon', '')}",
    ]
    if d.get("festnetz", "").strip():
        lines.append(f"TEL;TYPE=WORK,VOICE:{d['festnetz']}")

    adress_felder = ["strasse", "hausnummer", "plz", "ort", "land"]
    if any(d.get(f, "").strip() for f in adress_felder):
        strasse_nr = f"{d.get('strasse', '')} {d.get('hausnummer', '')}".strip()
        lines.append(f"ADR;TYPE=WORK:;;{strasse_nr};{d.get('ort', '')};;{d.get('plz', '')};{d.get('land', '')}")

    lines += [
        f"EMAIL;TYPE=WORK,INTERNET:{d.get('email', '')}",
        "END:VCARD",
    ]
    return "\r\n".join(lines)


def erstelle_qr(daten: Dict[str, str], 
                logo_pfad: Optional[Path], 
                fortschritt_cb: Optional[Callable[[int, str], None]] = None) -> Image.Image:
    if fortschritt_cb: fortschritt_cb(10, "Erstelle vCard …")

    vcard = erstelle_vcard(daten)

    qr = qrcode.QRCode(
        version=5,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=3,
    )
    qr.add_data(vcard)
    qr.make(fit=True)

    if fortschritt_cb: fortschritt_cb(40, "Rendere QR-Code …")

    qr_img = qr.make_image(fill_color=my_BLUE, back_color="white").convert("RGBA")

    if logo_pfad and logo_pfad.exists():
        if fortschritt_cb: fortschritt_cb(65, "Füge Logo ein …")
        try:
            logo = Image.open(logo_pfad).convert("RGBA")
            qr_w, qr_h = qr_img.size
            logo_size = int(qr_w * 0.25)
            logo = logo.resize((logo_size, logo_size), Image.Resampling.LANCZOS)
            pos_x = (qr_w - logo_size) // 2
            pos_y = (qr_h - logo_size) // 2
            qr_img.paste(logo, (pos_x, pos_y), logo)
        except Exception as e:
            logging.error(f"Fehler beim Verarbeiten des Logos: {e}")

    if fortschritt_cb: fortschritt_cb(90, "Finalisiere Bild …")

    return qr_img


# ─────────────────────────────────────────────
# Haupt-GUI
# ─────────────────────────────────────────────
class QRApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(False, False)
        self.configure(bg=BG_LIGHT)

        # State
        self.logo_pfad     = tk.StringVar()
        self.speicher_pfad = tk.StringVar()
        self.status_text   = tk.StringVar(value="Bereit.")
        self.profiles      = lade_profile()
        self.entries: Dict[str, tk.StringVar] = {}
        self.entry_widgets: Dict[str, tk.Entry] = {}
        self.placeholders: Dict[str, str] = {}
        self._preview_img: Optional[ImageTk.PhotoImage] = None
        self._full_qr_image: Optional[Image.Image] = None 

        self._build_ui()
        self._load_last_profile()

    def _build_ui(self):
        # -- Header --
        hdr = tk.Frame(self, bg=my_BLUE)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  vCard QR-Code Generator",
                 bg=my_BLUE, fg=WHITE,
                 font=("Segoe UI", 14, "bold"), pady=12).pack(side="left")

        # -- Haupt-Container --
        main = tk.Frame(self, bg=BG_LIGHT, padx=20, pady=16)
        main.pack(fill="both", expand=True)

        left = tk.Frame(main, bg=BG_LIGHT)
        left.grid(row=0, column=0, sticky="n", padx=(0, 20))

        right = tk.Frame(main, bg=BG_LIGHT)
        right.grid(row=0, column=1, sticky="n")

        # -- Profil-Sektion --
        self._section(left, "👤  Kontaktdaten")
        for key, label, placeholder in FELDER:
            var = tk.StringVar()
            self.entries[key] = var
            self.placeholders[key] = placeholder
            self.entry_widgets[key] = self._labeled_entry(left, label, var, placeholder)

        # -- Logo --
        self._section(left, "🖼️  Logo (optional)")
        logo_row = tk.Frame(left, bg=BG_LIGHT)
        logo_row.pack(fill="x", pady=4)
        tk.Entry(logo_row, textvariable=self.logo_pfad, width=30,
                 font=("Segoe UI", 9), fg=TEXT_COLOR, bg=WHITE,
                 relief="solid", bd=1).pack(side="left", ipady=4)
        self._btn(logo_row, "Durchsuchen", self._browse_logo,
                  bg=WHITE, fg=my_BLUE, bd=1, padx=8).pack(side="left", padx=(6, 0))

        # -- Speicherpfad --
        self._section(left, "💾  Ausgabe")
        out_row = tk.Frame(left, bg=BG_LIGHT)
        out_row.pack(fill="x", pady=4)
        tk.Entry(out_row, textvariable=self.speicher_pfad, width=30,
                 font=("Segoe UI", 9), fg=TEXT_COLOR, bg=WHITE,
                 relief="solid", bd=1).pack(side="left", ipady=4)
        self._btn(out_row, "Ordner wählen", self._browse_output,
                  bg=WHITE, fg=my_BLUE, bd=1, padx=8).pack(side="left", padx=(6, 0))

        # -- Profil-Verwaltung --
        self._section(left, "📋  Profile")
        prof_row = tk.Frame(left, bg=BG_LIGHT)
        prof_row.pack(fill="x", pady=4)

        self.profile_var = tk.StringVar()
        self.profile_box = ttk.Combobox(prof_row, textvariable=self.profile_var,
                                        width=22, state="readonly", font=("Segoe UI", 9))
        self.profile_box.pack(side="left", ipady=4)
        self._aktualisiere_profil_liste()

        self._btn(prof_row, "Laden",     self._lade_profil, bg=WHITE, fg=my_BLUE, bd=1, padx=6).pack(side="left", padx=(6, 0))
        self._btn(prof_row, "Speichern", self._speichere_profil, bg=ACCENT, fg=WHITE, padx=6).pack(side="left", padx=(4, 0))
        self._btn(prof_row, "🗑",        self._loesche_profil, bg=WHITE, fg=ERROR_COLOR, bd=1, padx=4).pack(side="left", padx=(4, 0))

        # -- Vorschau & Rechtsklick-Menü --
        self._section(right, "🔍  Vorschau (Rechtsklick = Kopieren)")
        preview_frame = tk.Frame(right, bg=WHITE, relief="solid", bd=1, width=250, height=250)
        preview_frame.pack_propagate(False)
        preview_frame.pack(pady=(4, 12))
        self.preview_label = tk.Label(preview_frame, bg=WHITE, text="Vorschau erscheint\nnach dem Generieren",
                                      fg="#999", font=("Segoe UI", 9), cursor="hand2")
        self.preview_label.pack(expand=True)

        # Rechtsklick Kontext-Menü
        self.context_menu = tk.Menu(self, tearoff=0, bg=WHITE, font=("Segoe UI", 9))
        self.context_menu.add_command(label="📋 Bild kopieren", command=self._copy_image_to_clipboard)
        
        self.preview_label.bind("<Button-3>", self._show_context_menu)
        self.preview_label.bind("<Button-2>", self._show_context_menu)

        # -- Aktions-Buttons --
        btn_frame = tk.Frame(right, bg=BG_LIGHT)
        btn_frame.pack(fill="x")
        self._btn(btn_frame, "🔄  Vorschau", lambda: self._generiere(nur_vorschau=True),
                  bg=WHITE, fg=my_BLUE, bd=1, padx=12).pack(fill="x", pady=(0, 6))
        self._btn(btn_frame, "✅  QR-Code speichern", lambda: self._generiere(nur_vorschau=False),
                  bg=my_BLUE, fg=WHITE, padx=12).pack(fill="x")

        # -- Fortschritt --
        prog_frame = tk.Frame(right, bg=BG_LIGHT)
        prog_frame.pack(fill="x", pady=(12, 0))
        self.progressbar = ttk.Progressbar(prog_frame, length=250, mode="determinate")
        self.progressbar.pack(fill="x")
        self.progress_label = tk.Label(prog_frame, textvariable=self.status_text, bg=BG_LIGHT, fg="#666", font=("Segoe UI", 8), anchor="w")
        self.progress_label.pack(fill="x", pady=(2, 0))

    # ── UI Hilfsmethoden ──────────────────────
    @staticmethod
    def _section(parent: tk.Widget, text: str) -> None:
        tk.Label(parent, text=text, bg=BG_LIGHT, fg=my_BLUE, font=("Segoe UI", 9, "bold"), anchor="w").pack(fill="x", pady=(12, 2))
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=(0, 6))

    def _labeled_entry(self, parent: tk.Widget, label: str, var: tk.StringVar, placeholder: str = "") -> tk.Entry:
        row = tk.Frame(parent, bg=BG_LIGHT)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=BG_LIGHT, fg=TEXT_COLOR, font=("Segoe UI", 9), width=20, anchor="w").pack(side="left")
        e = tk.Entry(row, textvariable=var, width=28, font=("Segoe UI", 9), bg=WHITE, relief="solid", bd=1)
        e.pack(side="left", ipady=4)
        if placeholder:
            var.set(placeholder)
            e.configure(fg="#666666")
            def on_focus_in(_, v=var, entry=e, ph=placeholder):
                if v.get() == ph:
                    v.set("")
                    entry.configure(fg=TEXT_COLOR)
            def on_focus_out(_, v=var, entry=e, ph=placeholder):
                if not v.get().strip():
                    v.set(ph)
                    entry.configure(fg="#666666")
            e.bind("<FocusIn>", on_focus_in)
            e.bind("<FocusOut>", on_focus_out)
        return e

    @staticmethod
    def _btn(parent: tk.Widget, text: str, cmd: Callable, bg: str = my_BLUE, fg: str = WHITE, bd: int = 0, padx: int = 16) -> tk.Button:
        return tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg, font=("Segoe UI", 9, "bold"),
                         relief="solid" if bd else "flat", bd=bd, padx=padx, cursor="hand2", activebackground=ACCENT, activeforeground=WHITE)

    # ── Zwischenablage Logik (KORRIGIERT) ─────
    def _show_context_menu(self, event) -> None:
        if self._full_qr_image:
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()

    def _copy_image_to_clipboard(self) -> None:
            """Kopiert das Bild robust in die Windows-Zwischenablage (64-Bit Safe)."""
            if not self._full_qr_image:
                return

            if sys.platform != "win32":
                messagebox.showinfo("Info", "Zwischenablage nur unter Windows verfügbar.")
                return

            try:
                import ctypes
                from ctypes import wintypes

                # Windows API Konstanten
                CF_DIB = 8
                GMEM_MOVEABLE = 0x0002
                
                # 1. Bilddaten vorbereiten (BMP Header entfernen für CF_DIB)
                output = io.BytesIO()
                self._full_qr_image.convert("RGB").save(output, "BMP")
                data = output.getvalue()[14:] # Header abschneiden
                output.close()
                
                # 2. Zugriff auf Windows-Bibliotheken
                user32 = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32
                
                # WICHTIG: Explizite Typendefinition für 64-Bit Python
                kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
                kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
                
                kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
                kernel32.GlobalLock.restype = wintypes.LPVOID
                
                kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
                kernel32.GlobalUnlock.restype = wintypes.BOOL
                
                kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
                kernel32.GlobalFree.restype = wintypes.HGLOBAL
                
                user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
                user32.SetClipboardData.restype = wintypes.HANDLE
                
                # 3. Speicher reservieren
                h_global = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
                if not h_global:
                    raise MemoryError("Konnte keinen Speicher reservieren (GlobalAlloc).")
                
                # 4. Speicher sperren und Daten kopieren
                target_ptr = kernel32.GlobalLock(h_global)
                if not target_ptr:
                    kernel32.GlobalFree(h_global)
                    raise MemoryError("Konnte Speicher nicht sperren (GlobalLock).")
                
                # Sicheres Kopieren mit ctypes.memmove
                ctypes.memmove(target_ptr, data, len(data))
                kernel32.GlobalUnlock(h_global)
                
                # 5. Zwischenablage öffnen (mit Retry)
                opened = False
                for _ in range(5):
                    if user32.OpenClipboard(None):
                        opened = True
                        break
                    time.sleep(0.1) # Kurz warten
                    
                if not opened:
                    kernel32.GlobalFree(h_global) # Wichtig: Aufräumen
                    raise OSError("Zwischenablage ist blockiert (OpenClipboard failed).")
                    
                user32.EmptyClipboard()
                
                # 6. Daten übergeben (System übernimmt Besitz bei Erfolg)
                if not user32.SetClipboardData(CF_DIB, h_global):
                    kernel32.GlobalFree(h_global) # Bei Fehler Speicher freigeben
                    user32.CloseClipboard()
                    raise OSError("SetClipboardData fehlgeschlagen.")
                    
                user32.CloseClipboard()
                self._set_status("📋 Bild in Zwischenablage kopiert!", SUCCESS)

            except Exception as e:
                logging.error(f"Fehler beim Kopieren: {e}")
                messagebox.showerror("Kopierfehler", f"Das Bild konnte nicht kopiert werden.\nGrund: {e}")




    # ── Interaktionen ─────────────────────────
    def _browse_logo(self) -> None:
        pfad = filedialog.askopenfilename(title="Logo auswählen", filetypes=[("Bilder", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Alle", "*.*")])
        if pfad: self.logo_pfad.set(pfad)

    def _browse_output(self) -> None:
        pfad = filedialog.askdirectory(title="Ausgabeordner wählen")
        if pfad: self.speicher_pfad.set(pfad)

    def _get_daten(self) -> Dict[str, str]:
        result = {}
        for k, v in self.entries.items():
            val = v.get().strip()
            result[k] = "" if val == self.placeholders.get(k, "") else val
        return result

    def _set_daten(self, daten: Dict[str, str]) -> None:
        for key, var in self.entries.items():
            value = daten.get(key, "")
            placeholder = self.placeholders.get(key, "")
            widget = self.entry_widgets.get(key)
            if value:
                var.set(value)
                if widget:
                    widget.configure(fg=TEXT_COLOR)
            else:
                var.set(placeholder)
                if widget:
                    widget.configure(fg="#666666" if placeholder else TEXT_COLOR)

    def _aktualisiere_profil_liste(self) -> None:
        namen = list(self.profiles.keys())
        self.profile_box["values"] = namen
        if namen and not self.profile_var.get(): self.profile_var.set(namen[-1])

    def _speichere_profil(self) -> None:
        daten = self._get_daten()
        name = f"{daten.get('vorname', '')} {daten.get('nachname', '')}".strip()
        if not name:
            messagebox.showwarning("Kein Name", "Bitte Vor- und Nachname eingeben.")
            return
        self.profiles[name] = daten
        speichere_profile(self.profiles)
        self._aktualisiere_profil_liste()
        self.profile_var.set(name)
        self._set_status(f"✅ Profil '{name}' gespeichert.", SUCCESS)

    def _lade_profil(self) -> None:
        name = self.profile_var.get()
        if name and name in self.profiles:
            self._set_daten(self.profiles[name])
            self._set_status(f"📂 Profil '{name}' geladen.", my_BLUE)

    def _loesche_profil(self) -> None:
        name = self.profile_var.get()
        if not name or name not in self.profiles: return
        if messagebox.askyesno("Profil löschen", f"Profil '{name}' wirklich löschen?"):
            del self.profiles[name]
            speichere_profile(self.profiles)
            self.profile_var.set("")
            self._aktualisiere_profil_liste()
            self._set_status(f"🗑 Profil '{name}' gelöscht.", ERROR_COLOR)

    def _load_last_profile(self) -> None:
        if self.profiles:
            last = list(self.profiles.keys())[-1]
            self._set_daten(self.profiles[last])
            self.profile_var.set(last)

    # ── Thread-sichere UI Updates ─────────────
    def _set_status(self, text: str, color: str = "#666") -> None:
        self.status_text.set(text)
        self.progress_label.configure(fg=color)

    def _thread_safe_fortschritt(self, wert: int, text: str = "") -> None:
        self.progressbar["value"] = wert
        if text: self._set_status(text)

    def _thread_safe_finish(self, success: bool, msg: str, preview_img: Optional[Image.Image] = None) -> None:
        if success:
            self._thread_safe_fortschritt(100, msg)
            if preview_img:
                self._full_qr_image = preview_img.copy()
                preview = preview_img.copy()
                preview.thumbnail((240, 240), Image.Resampling.LANCZOS)
                tk_img = ImageTk.PhotoImage(preview)
                self._preview_img = tk_img 
                self.preview_label.configure(image=tk_img, text="", cursor="hand2")
        else:
            self._thread_safe_fortschritt(0, msg)
            messagebox.showerror("Fehler", msg)

    # ── Generierung ───────────────────────────
    def _generiere(self, nur_vorschau: bool) -> None:
        if not QR_AVAILABLE:
            messagebox.showerror("Bibliothek fehlt", "Bitte installieren:\n  pip install qrcode[pil] Pillow")
            return

        daten = self._get_daten()
        fehler = [f for f in ["vorname", "nachname", "mobiltelefon", "email"] if not daten.get(f)]
        if fehler:
            messagebox.showwarning("Pflichtfelder", "Folgende Felder fehlen:\n" + "\n".join(f"• {f.capitalize()}" for f in fehler))
            return

        logo_pfad_str = self.logo_pfad.get().strip()
        logo = Path(logo_pfad_str) if logo_pfad_str else None
        outdir_str = self.speicher_pfad.get().strip()
        outdir = Path(outdir_str) if outdir_str else BASE_DIR

        self.progressbar["value"] = 0
        self._set_status("⏳ Generiere …")

        def task():
            try:
                img = erstelle_qr(daten, logo, lambda w, t: self.after(0, self._thread_safe_fortschritt, w, t))
                msg = "✅ Vorschau erstellt. (Rechtsklick zum Kopieren)"
                if not nur_vorschau:
                    dateiname = f"QR_{daten['vorname']}_{daten['nachname']}.png".replace(" ", "_")
                    ziel = outdir / dateiname
                    img.save(ziel)
                    msg = f"✅ Gespeichert: {ziel.name}"
                self.after(0, self._thread_safe_finish, True, msg, img)
            except Exception as exc:
                logging.error("Fehler bei der QR-Erstellung", exc_info=True)
                self.after(0, self._thread_safe_finish, False, f"❌ Fehler: {str(exc)}")

        threading.Thread(target=task, daemon=True).start()

if __name__ == "__main__":
    app = QRApp()
    app.mainloop()