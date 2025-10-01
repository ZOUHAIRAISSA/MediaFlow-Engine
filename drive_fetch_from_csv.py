# drive_fetch_from_csv.py
# -*- coding: utf-8 -*-

import os
import re
import io
import csv
import time
import zipfile
from pathlib import Path
from threading import Thread
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys  # <- ajouter

# --- Google API deps ---
# pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
try:
    PROJECT_ROOT = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
except Exception:
    PROJECT_ROOT = Path(__file__).resolve().parent  # racine du projet

# ---------- Helpers ----------
def log_to(app, msg: str):
    try:
        app.log.insert("end", msg + "\n")
        app.log.see("end")
        app.update_idletasks()
    except Exception:
        print(msg)

def _as_str(v) -> str:
    if isinstance(v, list):
        # garde le 1er élément non vide ou join si besoin
        flat = [("" if x is None else str(x)) for x in v]
        non_empty = [x for x in flat if x.strip()]
        v = non_empty[0] if non_empty else (flat[0] if flat else "")
    return ("" if v is None else str(v)).strip()

def read_csv_rows(csv_path: Path):
    with open(csv_path, "r", encoding="utf-8", errors="replace", newline="") as f:
        head = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(head, delimiters=",;")
            delim = dialect.delimiter
        except Exception:
            # heuristique simple
            delim = ";" if head.count(";") >= head.count(",") else ","
        rdr = csv.DictReader(f, delimiter=delim)
        for row in rdr:
            norm = {}
            for k, v in row.items():
                key = (k or "").lstrip("\ufeff").strip().lower()
                norm[key] = _as_str(v)
            yield norm

def extract_folder_id(link: str) -> str | None:
    s = str(link or "")
    m = re.search(r"/folders/([a-zA-Z0-9_-]{20,})", s)
    if m: return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]{20,})", s)
    if m: return m.group(1)
    m = re.search(r"[\w-]{25,}", s)
    return m.group(0) if m else None

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def sanitize_name(name: str) -> str:
    bad = '<>:"/\\|?*'
    out = "".join("_" if ch in bad else ch for ch in name)
    return out.rstrip(" .")[:150] or "item"

# ---------- Google Drive client ----------
class DriveClient:
    def __init__(self, creds_dir: Path | None = None):
        self.creds_dir = (creds_dir or PROJECT_ROOT)
        self.creds_dir.mkdir(parents=True, exist_ok=True)
        self.creds = None
        self.svc = None

    def authenticate(self):
        token_path = self.creds_dir / "token.json"
        creds = None
        if token_path.exists():
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(InstalledAppFlow.from_client_secrets_file)  # just to trigger import
                except Exception:
                    pass
            if not creds or not creds.valid:
                secrets_path = self.creds_dir / "credentials.json"
                if not secrets_path.exists():
                    raise RuntimeError(
                        f"credentials.json non trouvé dans le projet:\n{secrets_path}"
                    )
                flow = InstalledAppFlow.from_client_secrets_file(str(secrets_path), SCOPES)
                creds = flow.run_local_server(port=0)
            with open(token_path, "w", encoding="utf-8") as f:
                f.write(creds.to_json())
        self.creds = creds
        self.svc = build("drive", "v3", credentials=creds)

    def list_folder_files(self, folder_id: str):
        """
        Retourne une liste de dicts: id, name, mimeType, size
        """
        q = f"'{folder_id}' in parents and trashed=false"
        fields = "nextPageToken, files(id,name,mimeType,size)"
        items = []
        page_token = None
        while True:
            resp = self.svc.files().list(
                q=q, fields=fields, pageSize=1000, pageToken=page_token
            ).execute()
            items.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return items

    def download_file(self, file_id: str, dest_path: Path, mime_type: str | None, log_cb):
        """
        Télécharge un fichier binaire. Les Google Docs/Sheets/Slides seront ignorés ici.
        """
        # Google-native mimetypes
        if mime_type and mime_type.startswith("application/vnd.google-apps"):
            log_cb(f"  ⏭️ Ignoré (document Google natif): {dest_path.name}")
            return False

        request = self.svc.files().get_media(fileId=file_id)
        ensure_dir(dest_path.parent)
        fh = io.FileIO(dest_path, "wb")
        downloader = MediaIoBaseDownload(fh, request, chunksize=1024*1024)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                pct = int(status.progress() * 100)
                log_cb(f"    … {pct}% {dest_path.name}")
        fh.close()
        return True

    # NEW: récupérer le nom du dossier Drive
    def get_item_name(self, file_id: str) -> str:
        meta = self.svc.files().get(fileId=file_id, fields="id,name").execute()
        return meta.get("name", f"folder_{file_id[:6]}")

# ---------- Core job ----------
def download_from_csv(csv_path: Path, out_root: Path, creds_dir: Path | None = None, app_ui=None):
    """
    Lit CSV (file_download_csv_for_phtoshop), pour كل سطر:
      - يستخرج Drive Folder URL
      - يحمل جميع الملفات من ذلك المجلد إلى out_root/<sku_kyopa | sku_original>/
      - إذا كان الملف .zip → يفك الضغط
    """
    log = (lambda m: log_to(app_ui, m)) if app_ui else print

    dc = DriveClient(creds_dir or PROJECT_ROOT)
    log("🔐 Authentification Google Drive…")
    dc.authenticate()
    log("✅ Connecté à Google Drive.")

    rows = list(read_csv_rows(csv_path))
    if not rows:
        log("❌ CSV vide ou non valide.")
        return

    todo = []
    for r in rows:
        url = (r.get("drive folder url") or r.get("folder url") or r.get("url") or "").strip()
        if url:
            todo.append(r)
    if not todo:
        log("⚠️ Aucun « Drive Folder URL » dans le CSV.")
        return

    log(f"📥 {len(todo)} dossier(s) à traiter.")

    for idx, r in enumerate(todo, 1):
        # on n’utilise plus sku pour nommer le dossier local
        folder_url = (r.get("drive folder url") or r.get("folder url") or r.get("url") or "").strip()
        folder_id = extract_folder_id(folder_url)
        if not folder_id:
            log(f"#{idx} 🚫 Lien Drive invalide: {folder_url}")
            continue

        # Nom local = nom exact du dossier Drive
        try:
            drive_folder_name = dc.get_item_name(folder_id)
        except Exception as e:
            drive_folder_name = f"folder_{idx}"
            log(f"   ⚠️ Impossible de lire le nom du dossier Drive, fallback: {drive_folder_name} ({e})")

        subdir = sanitize_name(drive_folder_name)
        target_dir = out_root / subdir
        ensure_dir(target_dir)
        log(f"\n#{idx}/{len(todo)} 📂 Dossier: {drive_folder_name} → {target_dir.name}")

        try:
            files = dc.list_folder_files(folder_id)
        except Exception as e:
            log(f"   ❌ Erreur d’accès au dossier: {e}")
            continue

        if not files:
            log("   ⚠️ Aucun fichier dans ce dossier.")
            continue

        log(f"   Trouvé {len(files)} fichier(s). Téléchargement…")
        for f in files:
            name = sanitize_name(f.get("name", "file"))
            mime = f.get("mimeType", "")
            fid  = f.get("id")
            dest = target_dir / name
            try:
                ok = dc.download_file(fid, dest, mime, log)
                if ok and dest.suffix.lower() == ".zip":
                    # unzip then delete zip
                    try:
                        with zipfile.ZipFile(dest, 'r') as zf:
                            zf.extractall(target_dir)
                        dest.unlink(missing_ok=True)
                        log(f"    📦 Décompressé: {name}")
                    except Exception as e:
                        log(f"    ⚠️ Échec décompression {name}: {e}")
            except Exception as e:
                log(f"    ❌ Échec téléchargement {name}: {e}")

        # petite pause de courtoisie
        time.sleep(0.2)

    log("\n✅ Terminé: tous les dossiers traités.")

# ---------- Tkinter integration ----------
class DriveCSVDownloader(tk.Toplevel):
    def __init__(self, master, app_logger_widget=None):
        super().__init__(master)
        self.title("Télécharger dossiers Drive (CSV Photoshop)")
        self.geometry("640x360")
        self.resizable(True, True)
        self.app_ref = master
        self.log_widget = app_logger_widget

        self.csv_var = tk.StringVar()
        self.out_var = tk.StringVar()

        frm = ttk.LabelFrame(self, text="Paramètres")
        frm.pack(fill="x", padx=10, pady=8)

        self._row_path(frm, "CSV Photoshop:", self.csv_var, self.browse_csv)
        self._row_path(frm, "Dossier de sortie:", self.out_var, self.browse_out)

        hint = ttk.Label(self, text="Placez votre credentials.json dans la racine du projet.\n"
                                    "Le token.json sera créé automatiquement à côté.",
                         foreground="#555", justify="left")
        hint.pack(fill="x", padx=12)

        runbar = ttk.Frame(self); runbar.pack(fill="x", padx=10, pady=8)
        self.btn_start = ttk.Button(runbar, text="Télécharger maintenant", command=self.start_job)
        self.btn_start.pack(side="left")
        ttk.Button(runbar, text="Fermer", command=self.destroy).pack(side="left", padx=6)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=12, pady=(0,10))

    def _row_path(self, parent, label, var, cmd):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label, width=18).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Parcourir…", command=cmd).pack(side="left")

    def browse_csv(self):
        p = filedialog.askopenfilename(parent=self, title="Choisir le CSV",
                                       filetypes=[("CSV", "*.csv"), ("Tous les fichiers", "*.*")])
        if p: self.csv_var.set(p)

    def browse_out(self):
        p = filedialog.askdirectory(parent=self, title="Choisir le dossier de sortie")
        if p: self.out_var.set(p)

    def start_job(self):
        csvp = self.csv_var.get().strip()
        outp = self.out_var.get().strip()
        if not csvp or not outp:
            messagebox.showerror("Erreur", "Sélectionnez le CSV et le dossier de sortie.")
            return

        csv_path = Path(csvp)
        out_root = Path(outp)
        creds_dir = PROJECT_ROOT  # forcé: racine du projet

        def log_print(msg):
            if self.log_widget is not None:
                try:
                    self.log_widget.insert("end", msg + "\n")
                    self.log_widget.see("end")
                    self.log_widget.update_idletasks()
                except Exception:
                    print(msg)
            else:
                print(msg)

        def worker():
            try:
                download_from_csv(csv_path, out_root, creds_dir, app_ui=self.app_ref)
            except Exception as e:
                log_print(f"❌ Erreur: {e}")
            finally:
                self.progress.stop()
                self.btn_start.config(state="normal")

        self.btn_start.config(state="disabled")
        self.progress.start(10)
        Thread(target=worker, daemon=True).start()

# Hook à استدعاؤه من تطبيقك باش يضيف زرّ
def attach_drive_csv_downloader(app_instance, runbar_frame=None):
    """
    app_instance: كائن App الرئيسي ديالك (عندو .log)
    runbar_frame: الفريم اللي فيه الأزرار الرئيسية (اختياري).
                  إلا ما عطيتوهش، غادي نصايبو زرّ فوق-فوق في نافذة جديدة.
    """
    def open_window():
        win = DriveCSVDownloader(app_instance, app_logger_widget=getattr(app_instance, "log", None))
        win.transient(app_instance)
        win.grab_set()
        win.lift()
        win.focus_force()

    if runbar_frame is not None:
        ttk.Button(runbar_frame, text="Télécharger Drive (CSV)…",
                   command=open_window).pack(side="left", padx=6)
    else:
        # fallback: نافذة بسيطة فيها زرّ يفتح التول
        top = tk.Toplevel(app_instance)
        top.title("Intégration Drive")
        ttk.Button(top, text="Télécharger Drive (CSV)…", command=open_window).pack(padx=12, pady=12)
