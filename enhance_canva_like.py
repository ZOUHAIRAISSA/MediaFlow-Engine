import os
import subprocess
import shlex
import csv
from PIL import Image, ImageEnhance, ImageOps
import pillow_heif
from pathlib import Path
from threading import Thread
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ---------- ExifTool Functions (inspired from batchprocessor.py) ----------
def exiftool_bin() -> str:
    """Get exiftool executable path"""
    # Try to find exiftool.exe in the same directory as the script
    script_dir = Path(__file__).parent
    exiftool_path = script_dir / "exiftool.exe"
    
    if exiftool_path.exists():
        return str(exiftool_path)
    
    # Try parent directory (project root)
    parent_exiftool = script_dir.parent / "exiftool.exe"
    if parent_exiftool.exists():
        return str(parent_exiftool)
    
    # Fallback to system PATH
    return "exiftool.exe"

def build_exiftool_cmd_remove_metadata(out_path: Path) -> list[str]:
    """
    Build ExifTool command to remove ALL metadata from JPEG images.
    Inspired from batchprocessor.py build_exiftool_cmd_for_image function.
    """
    cmd = [
        exiftool_bin(),
        "-m", "-overwrite_original",
        "-charset", "UTF8", "-charset", "filename=UTF8",
        "-all=",        # Remove all metadata
        "-P"            # Preserve file modification time if possible
    ]
    cmd.append(str(out_path))
    return cmd

def remove_metadata_with_exiftool(out_path: Path, log_print=None) -> bool:
    """
    Remove all metadata from JPEG image using ExifTool.
    Returns True if successful, False otherwise.
    """
    if log_print is None:
        log_print = print
    
    try:
        et_cmd = build_exiftool_cmd_remove_metadata(out_path)
        log_print(f"[INFO] ExifTool cmd: {' '.join(shlex.quote(c) for c in et_cmd)}")
        
        res = subprocess.run(et_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        log_print("[OK] ExifTool: " + (res.stdout.strip() or "métadonnées supprimées."))
        return True
        
    except subprocess.CalledProcessError as e:
        log_print(f"[WARN] ExifTool a échoué: {e}")
        if e.stdout:
            log_print(f"[WARN] ExifTool stdout: {e.stdout}")
        return False
    except FileNotFoundError:
        log_print("[WARN] ExifTool non trouvé. Les métadonnées ne seront pas supprimées.")
        return False
    except Exception as e:
        log_print(f"[WARN] Erreur ExifTool: {e}")
        return False

# ---------- Helpers ----------
def apply_gamma(img, gamma=1.0):
    if gamma == 1.0:
        return img
    inv = 1.0 / max(gamma, 1e-6)
    lut = [min(255, int((i/255.0)**inv * 255 + 0.5)) for i in range(256)]
    return img.point(lut * (3 if img.mode == "RGB" else 1))

def wb_warm(img, r_gain=1.0, g_gain=1.0, b_gain=1.0):
    if img.mode != "RGB":
        return img.convert("RGB")
    r, g, b = img.split()
    r = r.point(lambda x: min(255, int(x * r_gain)))
    g = g.point(lambda x: min(255, int(x * g_gain)))
    b = b.point(lambda x: min(255, int(x * b_gain)))
    return Image.merge("RGB", (r, g, b))

def enhance_image_canva_like(image):
    """
    Canva-like natural preset: mild exposure/contrast/color/clarity.
    Values tuned for indoor rug shots.
    """
    # 1) احترام الـ EXIF orientation يكون قبل (نديرو خارج هاد الدالة)، هنا نفترض image مصححة
    img = image.convert("RGB")

    # 2) Autocontrast مع حماية 1% من الأطراف
    img = ImageOps.autocontrast(img, cutoff=1)

    # 3) Mid-tone lift (exposure خفيف)
    img = apply_gamma(img, gamma=0.98)  # <1 يرفع المتوسطات شويّة

    # 4) Color / Contrast / Brightness / Sharpness
    img = ImageEnhance.Brightness(img).enhance(1.03)
    img = ImageEnhance.Contrast(img).enhance(1.06)
    img = ImageEnhance.Color(img).enhance(1.08)      # vibrance تقريباً
    img = ImageEnhance.Sharpness(img).enhance(1.08)  # clarity خفيفة

    # 5) دفء خفيف
    img = wb_warm(img, r_gain=1.02, g_gain=1.00, b_gain=0.98)

    return img

def enhance_image_canva_custom(image, brightness=1.03, contrast=1.06, color=1.08, sharpness=1.08, gamma=0.98, r_gain=1.02, g_gain=1.00, b_gain=0.98):
    """
    Canva-like enhancement with customizable parameters.
    Based on enhance_image_canva_like but with adjustable values.
    """
    # 1) احترام الـ EXIF orientation يكون قبل (نديرو خارج هاد الدالة)، هنا نفترض image مصححة
    img = image.convert("RGB")

    # 2) Autocontrast مع حماية 1% من الأطراف
    img = ImageOps.autocontrast(img, cutoff=1)

    # 3) Mid-tone lift (exposure خفيف)
    img = apply_gamma(img, gamma=gamma)  # <1 يرفع المتوسطات شويّة

    # 4) Color / Contrast / Brightness / Sharpness
    img = ImageEnhance.Brightness(img).enhance(brightness)
    img = ImageEnhance.Contrast(img).enhance(contrast)
    img = ImageEnhance.Color(img).enhance(color)      # vibrance تقريباً
    img = ImageEnhance.Sharpness(img).enhance(sharpness)  # clarity خفيفة

    # 5) دفء خفيف
    img = wb_warm(img, r_gain=r_gain, g_gain=g_gain, b_gain=b_gain)

    return img

# ---------- Main convert ----------
def convert_and_enhance(input_folder, output_folder, resize_width=None, preset="none", canva_params=None):
    """
    Convert HEIC -> JPG, optionally enhance, keep metadata.
    Returns the path to the enhanced folder containing the images.
    
    Args:
        canva_params: dict with custom canva parameters (brightness, contrast, color, sharpness, gamma, r_gain, g_gain, b_gain)
    """
    # enhanced_folder = os.path.join(output_folder, os.path.basename(os.path.normpath(input_folder)))
    enhanced_folder = output_folder
    os.makedirs(enhanced_folder, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if not file_name.lower().endswith(".heic"):
            continue

        input_path = os.path.join(input_folder, file_name)
        output_name = os.path.splitext(file_name)[0] + ".jpg"
        enhanced_path = os.path.join(enhanced_folder, output_name)

        try:
            heif = pillow_heif.read_heif(input_path)
            img = Image.frombytes(heif.mode, heif.size, heif.data)

            # احترام orientation
            img = ImageOps.exif_transpose(img)

            # metadata (قد تكون غير موجودة)
            exif_bytes = heif.info.get("exif", None)
            icc_profile = heif.info.get("icc_profile", None)

            save_kwargs = {"format": "JPEG", "quality": 100, "subsampling": 0}
            if exif_bytes:
                save_kwargs["exif"] = exif_bytes
            if icc_profile:
                save_kwargs["icc_profile"] = icc_profile

            # Enhance
            if preset == "canva":
                if canva_params:
                    # Use custom parameters
                    out = enhance_image_canva_custom(img, **canva_params)
                else:
                    # Use default parameters
                    out = enhance_image_canva_like(img)
            elif preset == "none":
                out = img
            else:
                out = enhance_image_canva_like(img)  # default

            # Resize (اختياري)
            if resize_width:
                ar = out.height / out.width
                new_h = int(resize_width * ar)
                out = out.resize((resize_width, new_h), Image.Resampling.LANCZOS)

            out.save(enhanced_path, **save_kwargs)
            print(f"Enhanced and saved: {file_name} -> {output_name}")

        except Exception as e:
            print(f"Failed to process {file_name}: {e}")

    return enhanced_folder

# ---------- Metadata Removal Function ----------
def remove_metadata_from_folder(folder_path, log_print=None):
    """
    Remove metadata from all JPEG images in a folder and save them in the same folder.
   
    
    Args:
        folder_path: Path to folder containing JPEG images
        log_print: Optional logging function (default: print)
    
    Returns:
        dict: Summary of processing results
    """
    if log_print is None:
        log_print = print
    
    folder_path = Path(folder_path)
    if not folder_path.exists():
        log_print(f"❌ Error: Folder {folder_path} does not exist.")
        return {"success": False, "error": "Folder not found"}
    
    # Find all JPEG files
    jpeg_files = list(folder_path.glob("*.jpg")) + list(folder_path.glob("*.jpeg"))
    
    if not jpeg_files:
        log_print(f"❌ No JPEG files found in {folder_path}")
        return {"success": False, "error": "No JPEG files found"}
    
    log_print(f"🗑️ Starting metadata removal for {len(jpeg_files)} JPEG files...")
    log_print(f"📁 Folder: {folder_path}")
    log_print("")
    
    processed_count = 0
    failed_count = 0
    failed_files = []
    
    for jpeg_file in jpeg_files:
        try:
            log_print(f"🔄 Processing: {jpeg_file.name}")
            
            # Remove metadata using ExifTool
            success = remove_metadata_with_exiftool(jpeg_file, log_print)
            
            if success:
                log_print(f"✅ Metadata removed: {jpeg_file.name}")
                processed_count += 1
            else:
                log_print(f"⚠️ Failed to remove metadata: {jpeg_file.name}")
                failed_count += 1
                failed_files.append(jpeg_file.name)
                
        except Exception as e:
            log_print(f"❌ Error processing {jpeg_file.name}: {e}")
            failed_count += 1
            failed_files.append(jpeg_file.name)
    
    # Summary
    log_print(f"\n📊 Metadata Removal Summary:")
    log_print(f"✅ Successfully processed: {processed_count} files")
    log_print(f"❌ Failed: {failed_count} files")
    log_print(f"📁 Folder: {folder_path}")
    
    if failed_files:
        log_print(f"\n❌ Failed files:")
        for failed_file in failed_files:
            log_print(f"  - {failed_file}")
    
    return {
        "success": True,
        "processed_count": processed_count,
        "failed_count": failed_count,
        "failed_files": failed_files,
        "folder_path": str(folder_path)
    }

# ---------- CSV Functions (inspired from batchprocessor.py) ----------
def read_csv_data(csv_path):
    """Lit le CSV avec détection automatique du séparateur et retourne un dict: clé=sku original (minuscule)."""
    if not csv_path or not Path(csv_path).exists():
        return {}
    sku_data = {}
    try:
        with open(csv_path, 'r', encoding='utf-8', errors='replace', newline='') as f:
            sample = f.read(4096) or ""
            f.seek(0)
            # Détection du délimiteur
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;|\t")
                delim = dialect.delimiter
            except Exception:
                counts = {',': sample.count(','), ';': sample.count(';'), '\t': sample.count('\t'), '|': sample.count('|')}
                delim = max(counts, key=counts.get) if any(counts.values()) else ','
            reader = csv.DictReader(f, delimiter=delim)
            for row in reader:
                # normaliser les clés et valeurs
                norm = {}
                for k, v in row.items():
                    key = (k or "").lstrip("\ufeff").strip().lower()
                    val = "" if v is None else str(v).strip()
                    norm[key] = val
                sku_original = norm.get('sku original', '').strip().lower()
                if sku_original:
                    sku_data[sku_original] = norm
    except Exception as e:
        print(f"Erreur lors de la lecture du CSV: {e}")
    return sku_data

def build_exiftool_cmd_set_metadata(out_path: Path, title: str | None, tags: list[str] | None, rating: str | None) -> list[str]:
    """
    Build ExifTool command to set metadata for JPEG images.
    Inspired from batchprocessor.py build_exiftool_cmd_for_image function.
    """
    tags_list = [t.strip() for t in (tags or []) if t and t.strip()]
    tags_joined = ", ".join(tags_list)
    try:
        r = int(str(rating).strip()) if rating is not None else 5
    except Exception:
        r = 5
    pct = {1:1, 2:25, 3:50, 4:75, 5:99}.get(r, 99)

    cmd = [
        exiftool_bin(),
        "-m", "-overwrite_original",
        "-charset", "UTF8", "-charset", "filename=UTF8",
        "-sep", ", ",
        "-all=",        # Remove all metadata first
        "-P"            # Preserve file modification time if possible
    ]
    if title:
        cmd += [f"-XMP:Title={title}", f"-Xtra:Title={title}"]
    if tags_list:
        cmd += [f"-XMP-dc:Subject={tags_joined}", f"-Xtra:Keywords={tags_joined}"]
        for t in tags_list:
            cmd += [f"-XMP-dc:Subject={t}"]
    cmd += [f"-XMP-xmp:Rating={r}", f"-Xtra:Rating={pct}"]
    cmd.append(str(out_path))
    return cmd

def set_metadata_with_exiftool(out_path: Path, title: str | None, tags: list[str] | None, rating: str | None, log_print=None) -> bool:
    """
    Set metadata for JPEG image using ExifTool.
    Returns True if successful, False otherwise.
    """
    if log_print is None:
        log_print = print
    
    try:
        et_cmd = build_exiftool_cmd_set_metadata(out_path, title, tags, rating)
        log_print(f"[INFO] ExifTool cmd: {' '.join(shlex.quote(c) for c in et_cmd)}")
        
        res = subprocess.run(et_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        log_print("[OK] ExifTool: " + (res.stdout.strip() or "métadonnées écrites."))
        return True
        
    except subprocess.CalledProcessError as e:
        log_print(f"[WARN] ExifTool a échoué: {e}")
        if e.stdout:
            log_print(f"[WARN] ExifTool stdout: {e.stdout}")
        return False
    except FileNotFoundError:
        log_print("[WARN] ExifTool non trouvé. Les métadonnées ne seront pas écrites.")
        return False
    except Exception as e:
        log_print(f"[WARN] Erreur ExifTool: {e}")
        return False

def clean_filename(name):
    """Clean filename for Windows compatibility (inspired from batchprocessor.py)"""
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        name = name.replace(ch, "_")
    return (name or "image").rstrip(" .")[:150]

# ---------- GUI Interface (inspired from batchprocessor.py) ----------
class ImageEnhancerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🎨 Traitement d'images")
        self.geometry("900x800")
        self.configure(bg="#2b2b2b")
        self._setup_modern_style()
        self.resizable(True, True)
        
        # Variables
        self.input_folder_var = tk.StringVar()
        self.output_folder_var = tk.StringVar()
        self.csv_path_var = tk.StringVar()
        self.resize_width_var = tk.StringVar()
        self.preset_var = tk.StringVar(value="none")
        
        # Canva preset parameters (using existing values from code)
        self.brightness_var = tk.DoubleVar(value=1.03)
        self.contrast_var = tk.DoubleVar(value=1.06)
        self.color_var = tk.DoubleVar(value=1.08)
        self.sharpness_var = tk.DoubleVar(value=1.08)
        self.gamma_var = tk.DoubleVar(value=0.98)
        self.r_gain_var = tk.DoubleVar(value=1.02)
        self.g_gain_var = tk.DoubleVar(value=1.00)
        self.b_gain_var = tk.DoubleVar(value=0.98)
        
        # Create GUI
        self._create_widgets()
        
        # Processing state
        self._stop_requested = False
        self._worker = None
        
    def _setup_modern_style(self):
        """Setup modern dark theme (inspired from batchprocessor.py)"""
        try:
            style = ttk.Style()
            style.theme_use('clam')
            # Base colors
            bg = '#2b2b2b'
            fg = '#ffffff'
            btn_bg = '#4a90e2'
            btn_active = '#357abd'
            success_bg = '#28a745'
            success_active = '#218838'
            danger_bg = '#dc3545'
            danger_active = '#c82333'
            
            # Frames and labels
            style.configure('TFrame', background=bg)
            style.configure('TLabelframe', background=bg, foreground=fg)
            style.configure('TLabelframe.Label', background=bg, foreground=fg)
            style.configure('TLabel', background=bg, foreground=fg)
            
            # Buttons
            style.configure('TButton', background=btn_bg, foreground=fg)
            style.map('TButton', background=[('active', btn_active)])
            style.configure('Success.TButton', background=success_bg, foreground=fg)
            style.map('Success.TButton', background=[('active', success_active)])
            style.configure('Danger.TButton', background=danger_bg, foreground=fg)
            style.map('Danger.TButton', background=[('active', danger_active)])
            
            # Entries
            style.configure('TEntry', fieldbackground='#3a3a3a', foreground=fg)
        except Exception:
            pass
    
    def _create_widgets(self):
        """Create all GUI widgets"""
        # Main title
        title_frame = ttk.Frame(self)
        title_frame.pack(fill="x", padx=10, pady=10)
        title_label = ttk.Label(title_frame, text="🎨 Traitement d'images", 
                               font=("Arial", 16, "bold"))
        title_label.pack()
        
        # Paths section
        paths_frame = ttk.LabelFrame(self, text="📂 Chemins")
        paths_frame.pack(fill="x", padx=10, pady=8)
        
        self._row_path(paths_frame, "Dossier parent (sous-dossiers):", self.input_folder_var, self.browse_input)
        self._row_path(paths_frame, "Dossier de sortie:", self.output_folder_var, self.browse_output)
        self._row_path(paths_frame, "Fichier CSV:", self.csv_path_var, self.browse_csv)
        
        # Options section
        options_frame = ttk.LabelFrame(self, text="⚙️ Options")
        options_frame.pack(fill="x", padx=10, pady=8)
        
        self._row_entry(options_frame, "Largeur de redimensionnement:", self.resize_width_var)
        
        # Preset section
        preset_frame = ttk.LabelFrame(self, text="🎨 Presets d'amélioration")
        preset_frame.pack(fill="x", padx=10, pady=8)
        
        # Preset selection
        preset_row = ttk.Frame(preset_frame)
        preset_row.pack(fill="x", padx=8, pady=4)
        
        ttk.Label(preset_row, text="Preset:", width=15).pack(side="left")
        preset_combo = ttk.Combobox(preset_row, textvariable=self.preset_var, 
                                   values=["none", "canva"], state="readonly", width=15)
        preset_combo.pack(side="left", padx=6)
        preset_combo.bind("<<ComboboxSelected>>", self._on_preset_change)
        
        # Canva preset parameters (initially hidden)
        self.canva_params_frame = ttk.LabelFrame(preset_frame, text="Paramètres Canva")
        
        # Create parameter rows
        self._row_scale(self.canva_params_frame, "Luminosité:", self.brightness_var, 0.5, 2.0, 0.01)
        self._row_scale(self.canva_params_frame, "Contraste:", self.contrast_var, 0.5, 2.0, 0.01)
        self._row_scale(self.canva_params_frame, "Couleur:", self.color_var, 0.5, 2.0, 0.01)
        self._row_scale(self.canva_params_frame, "Netteté:", self.sharpness_var, 0.5, 2.0, 0.01)
        self._row_scale(self.canva_params_frame, "Gamma:", self.gamma_var, 0.5, 1.5, 0.01)
        self._row_scale(self.canva_params_frame, "Gain Rouge:", self.r_gain_var, 0.5, 1.5, 0.01)
        self._row_scale(self.canva_params_frame, "Gain Vert:", self.g_gain_var, 0.5, 1.5, 0.01)
        self._row_scale(self.canva_params_frame, "Gain Bleu:", self.b_gain_var, 0.5, 1.5, 0.01)
        
        
        # Control buttons
        control_frame = ttk.Frame(self)
        control_frame.pack(fill="x", padx=10, pady=8)
        
        self.start_btn = ttk.Button(control_frame, text="🚀 Démarrer le traitement", 
                                   command=self.start_processing, style='Success.TButton')
        self.start_btn.pack(side="left", padx=5)
        
        self.stop_btn = ttk.Button(control_frame, text="⏹️ Arrêter", 
                                  command=self.stop_processing, state="disabled", style='Danger.TButton')
        self.stop_btn.pack(side="left", padx=5)
        
        # Log section
        log_frame = ttk.LabelFrame(self, text="📋 Journal")
        log_frame.pack(fill="both", expand=True, padx=10, pady=8)
        
        self.log_text = tk.Text(log_frame, height=15, bg="#1f1f1f", fg="#eaeaea", 
                               insertbackground="#eaeaea", borderwidth=0, highlightthickness=0)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Now that log_text is created, we can call _on_preset_change
        self._on_preset_change()
        
        # Initial message
        self._append("🎨 Bienvenue dans l'Image Enhancer avec intégration CSV!")
        self._append("📂 Sélectionnez le dossier parent, le dossier de sortie et le fichier CSV.")
        self._append("⚙️ Optionnel: spécifiez une largeur de redimensionnement.")
        self._append("🚀 Cliquez sur 'Démarrer le traitement' pour commencer.\n")
    
    def _row_path(self, parent, label, var, command):
        """Create a row with label, entry and browse button"""
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        
        ttk.Label(row, text=label, width=25).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Parcourir…", command=command).pack(side="left")
    
    def _row_entry(self, parent, label, var, placeholder=""):
        """Create a row with label and entry"""
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        
        ttk.Label(row, text=label, width=25).pack(side="left")
        entry = ttk.Entry(row, textvariable=var)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        if placeholder:
            entry.insert(0, placeholder)
        return entry
    
    def _row_scale(self, parent, label, var, from_, to_, resolution=0.01):
        """Create a row with label and scale (inspired from batchprocessor.py)"""
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=3)
        
        ttk.Label(row, text=label, width=15).pack(side="left")
        val_lbl = ttk.Label(row, text=str(var.get()))
        val_lbl.pack(side="right")
        
        scale = ttk.Scale(
            row, from_=from_, to=to_, variable=var, orient="horizontal",
            command=lambda v: val_lbl.config(
                text=f"{float(v):.2f}" if resolution != 1 else f"{int(float(v))}"
            )
        )
        scale.pack(side="left", fill="x", expand=True, padx=6)
        return scale
    
    def _on_preset_change(self, event=None):
        """Handle preset selection change"""
        preset = self.preset_var.get()
        if preset == "canva":
            self.canva_params_frame.pack(fill="x", padx=8, pady=4)
            self._append("🎨 Preset 'canva' sélectionné - Paramètres ajustables affichés")
        else:
            self.canva_params_frame.pack_forget()
            self._append("🎨 Preset 'none' sélectionné - Aucune amélioration appliquée")
    
    def _append(self, text):
        """Append text to log"""
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.update_idletasks()
    
    def browse_input(self):
        """Browse for input parent folder"""
        folder = filedialog.askdirectory(title="Choisir le dossier parent (contenant les sous-dossiers)")
        if folder:
            self.input_folder_var.set(folder)
            self._append(f"📂 Dossier parent sélectionné: {folder}")
    
    def browse_output(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if folder:
            self.output_folder_var.set(folder)
            self._append(f"📁 Dossier de sortie sélectionné: {folder}")
    
    def browse_csv(self):
        """Browse for CSV file"""
        file_path = filedialog.askopenfilename(
            title="Choisir le fichier CSV",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.csv_path_var.set(file_path)
            self._append(f"📄 Fichier CSV sélectionné: {file_path}")
            
            # Load and preview CSV data
            csv_data = read_csv_data(file_path)
            if csv_data:
                self._append(f"✅ CSV chargé: {len(csv_data)} entrées trouvées")
                self._append("📋 Aperçu des données CSV:")
                sample = list(csv_data.items())[:3]  # Show first 3 entries
                for sku_orig, row in sample:
                    sku_kyopa = row.get('sku kyopa', '')
                    title = row.get('title', '')
                    tags = row.get('tags', '')
                    self._append(f"  • '{sku_orig}' → SKU: '{sku_kyopa}'")
                    self._append(f"    Titre: '{title}'")
                    self._append(f"    Tags: '{tags}'")
            else:
                self._append("⚠️ Aucune donnée valide trouvée dans le CSV")
    
    def start_processing(self):
        """Start the image processing"""
        input_folder = self.input_folder_var.get().strip()
        output_folder = self.output_folder_var.get().strip()
        csv_path = self.csv_path_var.get().strip()
        resize_width_str = self.resize_width_var.get().strip()
        
        # Validation
        if not input_folder or not output_folder or not csv_path:
            messagebox.showerror("Erreur", "Veuillez sélectionner tous les chemins requis:\n• Dossier parent\n• Dossier de sortie\n• Fichier CSV")
            return
        
        if not os.path.isdir(input_folder):
            messagebox.showerror("Erreur", f"Le dossier parent '{input_folder}' n'existe pas ou n'est pas un dossier.")
            return
        
        if not Path(csv_path).exists():
            messagebox.showerror("Erreur", f"Le fichier CSV '{csv_path}' n'existe pas.")
            return
        
        # Parse resize width
        resize_width = None
        if resize_width_str:
            try:
                resize_width = int(resize_width_str)
            except ValueError:
                messagebox.showerror("Erreur", "La largeur de redimensionnement doit être un nombre entier.")
                return
        
        # Update UI
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._stop_requested = False
        
        self._append("\n" + "="*60)
        self._append("🚀 DÉMARRAGE DU TRAITEMENT")
        self._append("="*60)
        self._append(f"📂 Dossier parent: {input_folder}")
        self._append(f"📁 Dossier de sortie: {output_folder}")
        self._append(f"📄 Fichier CSV: {csv_path}")
        if resize_width:
            self._append(f"📏 Largeur de redimensionnement: {resize_width}px")
        self._append("")
        
        # Get preset selection
        preset = self.preset_var.get()
        
        # Start processing in background thread
        def worker():
            try:
                self._run_processing(input_folder, output_folder, csv_path, resize_width, preset)
            except Exception as e:
                self._append(f"❌ Erreur inattendue: {e}")
            finally:
                self.start_btn.config(state="normal")
                self.stop_btn.config(state="disabled")
                self._append("\n🎉 Traitement terminé!")
        
        self._worker = Thread(target=worker, daemon=True)
        self._worker.start()
    
    def _run_processing(self, input_folder, output_folder, csv_path, resize_width, preset):
        """Run the actual processing (adapted from main block)"""
        # Load CSV data
        self._append("📄 Chargement des données CSV...")
        csv_data = read_csv_data(csv_path)
        if not csv_data:
            self._append("❌ Erreur: Aucune donnée valide trouvée dans le fichier CSV")
            return
        
        self._append(f"✅ CSV chargé: {len(csv_data)} entrées trouvées")
        self._append(f"🎨 Preset sélectionné: {preset}")
        
        if preset == "canva":
            self._append("📊 Paramètres Canva:")
            self._append(f"  - Luminosité: {self.brightness_var.get():.2f}")
            self._append(f"  - Contraste: {self.contrast_var.get():.2f}")
            self._append(f"  - Couleur: {self.color_var.get():.2f}")
            self._append(f"  - Netteté: {self.sharpness_var.get():.2f}")
            self._append(f"  - Gamma: {self.gamma_var.get():.2f}")
            self._append(f"  - Gain Rouge: {self.r_gain_var.get():.2f}")
            self._append(f"  - Gain Vert: {self.g_gain_var.get():.2f}")
            self._append(f"  - Gain Bleu: {self.b_gain_var.get():.2f}")
        
        input_path = Path(input_folder)
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Find all subfolders
        subfolders = [f for f in input_path.iterdir() if f.is_dir()]
        
        if not subfolders:
            self._append("❌ Aucun sous-dossier trouvé dans le dossier parent!")
            return
        
        self._append(f"📂 {len(subfolders)} sous-dossiers trouvés:")
        for i, subfolder in enumerate(subfolders, 1):
            self._append(f"  {i}. {subfolder.name}")
        self._append("")
        
        # Process each subfolder
        total_processed = 0
        total_failed = 0
        processed_folders = []
        failed_folders = []
        csv_matched_folders = []
        csv_unmatched_folders = []
        
        for i, subfolder in enumerate(subfolders, 1):
            if self._stop_requested:
                self._append("⏹️ Arrêt demandé par l'utilisateur")
                break
                
            self._append(f"🔄 Traitement du sous-dossier {i}/{len(subfolders)}: {subfolder.name}")
            
            # Look for CSV match
            subfolder_name_lower = subfolder.name.strip().lower()
            csv_match = csv_data.get(subfolder_name_lower)
            
            if csv_match:
                self._append(f"✅ Correspondance CSV trouvée pour '{subfolder.name}'")
                csv_title = csv_match.get('title', '').strip()
                csv_sku_kyopa = csv_match.get('sku kyopa', '').strip()
                csv_tags = csv_match.get('tags', '').strip()
                
                self._append(f"📋 Données CSV:")
                self._append(f"  - Titre: '{csv_title}'")
                self._append(f"  - SKU Kyopa: '{csv_sku_kyopa}'")
                self._append(f"  - Tags: '{csv_tags}'")
                
                output_folder_name = csv_sku_kyopa if csv_sku_kyopa else subfolder.name
                csv_matched_folders.append(subfolder.name)
            else:
                self._append(f"⚠️ Aucune correspondance CSV trouvée pour '{subfolder.name}' - utilisation du nom original")
                output_folder_name = subfolder.name
                csv_title = ""
                csv_tags = ""
                csv_unmatched_folders.append(subfolder.name)
            
            # Create output subfolder
            output_subfolder = output_path / output_folder_name
            output_subfolder.mkdir(parents=True, exist_ok=True)
            self._append(f"📁 Dossier de sortie: {output_subfolder}")
            
            # Step 1: Convert and enhance
            self._append("🔄 Étape 1: Conversion et amélioration des images...")
            
            # Prepare canva parameters if preset is canva
            canva_params = None
            if preset == "canva":
                canva_params = {
                    'brightness': self.brightness_var.get(),
                    'contrast': self.contrast_var.get(),
                    'color': self.color_var.get(),
                    'sharpness': self.sharpness_var.get(),
                    'gamma': self.gamma_var.get(),
                    'r_gain': self.r_gain_var.get(),
                    'g_gain': self.g_gain_var.get(),
                    'b_gain': self.b_gain_var.get()
                }
            
            enhanced_folder = convert_and_enhance(str(subfolder), str(output_subfolder), resize_width, preset=preset, canva_params=canva_params)
            
            if enhanced_folder and Path(enhanced_folder).exists():
                self._append(f"✅ Étape 1 terminée pour '{subfolder.name}'!")
                
                # Step 2: Remove metadata
                self._append("🗑️ Étape 2: Suppression des métadonnées...")
                result = remove_metadata_from_folder(enhanced_folder, self._append)
                
                if result["success"]:
                    self._append(f"✅ Suppression des métadonnées terminée!")
                    self._append(f"🗑️ Métadonnées supprimées de {result['processed_count']} fichiers")
                    
                    # Step 3: Apply CSV metadata and rename
                    if csv_match and csv_title:
                        self._append("📄 Étape 3: Application des métadonnées CSV et renommage...")
                        
                        # Process tags
                        if csv_tags:
                            tags = [t.strip() for t in (csv_tags.split(",") if ',' in csv_tags else csv_tags.split()) if t.strip()]
                        else:
                            tags = []
                        
                        # Clean title for filename
                        safe_title = clean_filename(csv_title)
                        
                        # Find all JPEG files and rename them
                        jpeg_files = list(Path(enhanced_folder).glob("*.jpg")) + list(Path(enhanced_folder).glob("*.jpeg"))
                        
                        processed_images = 0
                        for j, jpeg_file in enumerate(jpeg_files, 1):
                            try:
                                # Create new filename with sequential number
                                new_filename = f"{safe_title}_{j}.jpg"
                                new_path = Path(enhanced_folder) / new_filename
                                
                                # Rename the file
                                jpeg_file.rename(new_path)
                                self._append(f"📝 Renommé: {jpeg_file.name} -> {new_filename}")
                                
                                # Set metadata using ExifTool
                                success = set_metadata_with_exiftool(new_path, csv_title, tags, "5", self._append)
                                
                                if success:
                                    self._append(f"✅ Métadonnées appliquées: {new_filename}")
                                    processed_images += 1
                                else:
                                    self._append(f"⚠️ Échec de l'application des métadonnées: {new_filename}")
                                    
                            except Exception as e:
                                self._append(f"❌ Erreur lors du traitement de {jpeg_file.name}: {e}")
                        
                        self._append(f"✅ Métadonnées CSV appliquées à {processed_images} images")
                    else:
                        self._append(f"ℹ️ Aucune donnée CSV à appliquer pour '{subfolder.name}'")
                    
                    total_processed += result['processed_count']
                    processed_folders.append(subfolder.name)
                else:
                    self._append(f"⚠️ Problème avec la suppression des métadonnées pour '{subfolder.name}', mais la conversion a réussi.")
                    total_failed += 1
                    failed_folders.append(subfolder.name)
            else:
                self._append(f"❌ Étape 1 échouée pour '{subfolder.name}'!")
                total_failed += 1
                failed_folders.append(subfolder.name)
            
            self._append("-" * 40)
        
        # Final summary
        self._append("\n🎉 TRAITEMENT PAR LOTS TERMINÉ!")
        self._append("📊 Résumé:")
        self._append(f"  📂 Total des sous-dossiers traités: {len(subfolders)}")
        self._append(f"  ✅ Traités avec succès: {len(processed_folders)}")
        self._append(f"  ❌ Échoués: {len(failed_folders)}")
        self._append(f"  🖼️ Total des images traitées: {total_processed}")
        self._append(f"  📄 Dossiers avec correspondance CSV: {len(csv_matched_folders)}")
        self._append(f"  ⚠️ Dossiers sans correspondance CSV: {len(csv_unmatched_folders)}")
        
        if csv_matched_folders:
            self._append(f"\n✅ Dossiers avec correspondance CSV:")
            for folder in csv_matched_folders:
                self._append(f"  - {folder}")
        
        if csv_unmatched_folders:
            self._append(f"\n⚠️ Dossiers sans correspondance CSV:")
            for folder in csv_unmatched_folders:
                self._append(f"  - {folder}")
    
    def stop_processing(self):
        """Stop the processing"""
        self._stop_requested = True
        self._append("⏹️ Arrêt demandé... Le traitement s'arrêtera après le fichier en cours.")
        messagebox.showinfo("Info", "Le traitement s'arrêtera après avoir terminé le fichier en cours.")

if __name__ == "__main__":
    # Launch GUI interface
    app = ImageEnhancerApp()
    app.mainloop()
