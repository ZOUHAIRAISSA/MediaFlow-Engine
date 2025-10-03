#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from drive_fetch_from_csv import attach_drive_csv_downloader
import argparse
import os
import sys
import shutil
import tempfile
import shlex
import subprocess
import csv
from pathlib import Path
from threading import Thread
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
# import shutil  # exiftool lookup

# PyWin32 (pour écrire dans Propriétés Windows)
try:
    from win32com.propsys import propsys, pscon
    import pythoncom
    HAVE_PYWIN32 = True
except Exception:
    HAVE_PYWIN32 = False

# Ajout: import Mutagen (optionnel)
try:
    from mutagen.mp4 import MP4, MP4FreeForm
except Exception:
    MP4 = None
    MP4FreeForm = None

# --- Images (HEIC/JPEG) ---
from PIL import Image
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()   # يخلي PIL يفتح HEIC
    HAVE_HEIF = True
except Exception:
    HAVE_HEIF = False




# =======================
#  Core processing logic
# =======================

VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".mkv", ".avi", ".wmv", ".flv", ".webm"}
IMAGE_EXTS = {".heic", ".jpg", ".jpeg"}  



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

def infer_tags_from_path(src: Path, root: Path) -> list[str]:
    parts = list(src.relative_to(root).parts)
    if len(parts) > 1:
        parts = parts[:-1]
    tags = [p.strip() for p in parts if p.strip()]
    seen = set()
    uniq = []
    for t in tags:
        tl = t.lower()
        if tl not in seen:
            uniq.append(t)
            seen.add(tl)
    return uniq

def safe_mkdirs(p: Path):
    p.parent.mkdir(parents=True, exist_ok=True)


def _app_cache_dir() -> Path:
    """
    فولدر كاش قابل للكتابة للمستخدم الحالي:
    %LOCALAPPDATA%\BatchVideoProcessor\tools
    """
    base = Path(os.getenv("LOCALAPPDATA", tempfile.gettempdir())) / "BatchVideoProcessor" / "tools"
    base.mkdir(parents=True, exist_ok=True)
    return base

def _bundle_path(name: str) -> Path | None:
    """
    كنجيب المسار إلى الملف داخل الباندل (_MEIPASS) إذا كان موجود.
    وإلا كنقلّب عليه جنب السكريبت فالسورس.
    """
    mei = getattr(sys, "_MEIPASS", None)
    if mei:
        p = Path(mei) / name
        return p if p.exists() else None
    p = Path(__file__).parent / name
    return p if p.exists() else None

def _ensure_tool(name: str) -> str:
    """
    كنرجّع مسار نهائي للأداة (exiftool.exe أو ffmpeg.exe).
    منطق الأولويات:
      - إذا كنّا غير مجمَّدين (not frozen): وخا كان الملف جنب السكريبت، استعملو مباشرة.
      - إذا كنّا مجمَّدين (--onefile): ننسخو من _MEIPASS إلى كاش قابل للكتابة ثم نستعملو من الكاش.
      - وإلا fallback على PATH.
    """
    frozen = hasattr(sys, "_MEIPASS")

    # في التطوير (not frozen): استعمل الملف المحلي جنب السكريبت إذا كان موجود
    local_path = Path(__file__).parent / name
    if local_path.exists() and not frozen:
        if name.lower().startswith("exiftool"):
            local_exif_dir = Path(__file__).parent / "exiftool_files"
            if local_exif_dir.exists():
                os.environ.setdefault("EXIFTOOL_HOME", str(local_exif_dir))
        return str(local_path)

    # في الـ onefile: خُد من الباندل ونسخو للكاش
    src = _bundle_path(name)
    if src and src.exists():
        dst_dir = _app_cache_dir()
        dst = dst_dir / name
        try:
            # نسخ إذا ما كايناش أو الحجم تبدّل (مؤشر بدائي للتحديث)
            if (not dst.exists()) or (os.path.getsize(dst) != os.path.getsize(src)):
                shutil.copy2(src, dst)
        except Exception:
            # إلى فشل النسخ لأي سبب، رجّع الاسم وخلي PATH يقرره
            return name

        # EXIFTOOL_HOME خاص يكون فولدر قابل للكتابة
        if name.lower().startswith("exiftool"):
            src_exif_dir = (Path(src).parent / "exiftool_files")
            if src_exif_dir.exists():
                target_exif_dir = dst_dir / "exiftool_files"
                if not target_exif_dir.exists():
                    try:
                        shutil.copytree(src_exif_dir, target_exif_dir)
                    except Exception:
                        pass
                os.environ["EXIFTOOL_HOME"] = str(target_exif_dir)
            else:
                os.environ["EXIFTOOL_HOME"] = str(dst_dir)
        return str(dst)

    # آخر حل: PATH
    return name

#add fonction to convert heic to jpg
def convert_heic_to_jpg(heic_path: Path, jpg_out: Path | None = None, quality: int = 100) -> Path:
    """
    HEIC -> JPG باستعمال pillow-heif + Pillow
    quality 0..100 (95 جيد جداً). نقدر ندير 100 باش ننقص الفقدان لأدنى حد.
    """
    if not heic_path.exists():
        raise RuntimeError(f"Fichier source introuvable: {heic_path}")
    if not HAVE_HEIF:
        raise RuntimeError("pillow-heif غير متاح. ثبّت: pip install pillow-heif")

    if jpg_out is None:
        jpg_out = heic_path.with_suffix(".jpg")

    try:
        with Image.open(heic_path) as im:
            rgb = im.convert("RGB")
            # subsampling=0 يحافظ على جودة كبرى (4:4:4)
            rgb.save(jpg_out, format="JPEG", quality=quality, subsampling=0, optimize=True)
    except Exception as e:
        raise RuntimeError(f"Erreur conversion HEIC: {e}")

    if not jpg_out.exists():
        raise RuntimeError("La conversion n'a pas généré de fichier")
    return jpg_out


def exiftool_bin() -> str:
    return _ensure_tool("exiftool.exe")

def ffmpeg_bin() -> str:
    return _ensure_tool("ffmpeg.exe")

def set_win_explorer_props_mp4(file_path: str, title: str | None, tags: list[str], stars: int, log_print):
    """
    Écrit System.Title, System.Keywords et System.Rating via IPropertyStore.
    Nécessite pywin32. Fonctionne même si ExifTool a échoué.
    """
    try:
        if not HAVE_PYWIN32:
            log_print("[WINPROPS] pywin32 non disponible. py -m pip install pywin32")
            return False

        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        flags = getattr(pscon, "GPS_READWRITE", 0x2)
        store = propsys.SHGetPropertyStoreFromParsingName(
            os.fspath(file_path), None, flags, propsys.IID_IPropertyStore
        )

        # Clés
        key_title    = getattr(pscon, "PKEY_Title",    None) or propsys.PSGetPropertyKeyFromName("System.Title")
        key_keywords = getattr(pscon, "PKEY_Keywords", None) or propsys.PSGetPropertyKeyFromName("System.Keywords")
        key_rating   = getattr(pscon, "PKEY_Rating",   None) or propsys.PSGetPropertyKeyFromName("System.Rating")
        try:
            key_media_title = propsys.PSGetPropertyKeyFromName("System.Media.Title")
        except Exception:
            key_media_title = None

        # Title
        if title:
            store.SetValue(key_title, propsys.PROPVARIANTType(str(title)))
            if key_media_title:
                store.SetValue(key_media_title, propsys.PROPVARIANTType(str(title)))

        # Keywords (liste)
        clean = [t.strip() for t in (tags or []) if t and t.strip()]
        try:
            pv_keywords = propsys.PROPVARIANTType(clean) if clean else propsys.PROPVARIANTType(None)
        except Exception:
            # fallback chaîne unique
            pv_keywords = propsys.PROPVARIANTType(", ".join(clean)) if clean else propsys.PROPVARIANTType(None)
        store.SetValue(key_keywords, pv_keywords)

        # Rating: map 1..5 -> 1/25/50/75/99
        s = max(1, min(int(stars or 5), 5))
        rating99 = {1: 1, 2: 25, 3: 50, 4: 75, 5: 99}[s]
        store.SetValue(key_rating, propsys.PROPVARIANTType(int(rating99)))

        store.Commit()
        try:
            os.utime(file_path, None)
        except Exception:
            pass

        log_print(f"[WINPROPS] OK: Title='{title}' Keywords={clean} Rating={s} → {file_path}")
        return True
    except Exception as e:
        log_print(f"[WINPROPS] Echec: {e}")
        return False

# Stub Mutagen (évite l’erreur "not defined")
def write_mp4_metadata_mutagen(path: Path, title: str | None, tags: list[str] | None, rating: int, log_print):
    if MP4 is None:
        raise RuntimeError("Mutagen non disponible")
    m = MP4(str(path))
    if title:
        m["\xa9nam"] = [str(title)]
    tag_text = ", ".join([t for t in (tags or []) if t.strip()])
    if tag_text:
        m["\xa9cmt"] = [tag_text]
    # Note: Mutagen n’écrit pas les balises Windows Explorer (Xtra/Property System)
    m.save()
    log_print("[OK] Mutagen: titre/comment MP4 écrits.")

def build_ffmpeg_cmd(
    inp: Path,
    out: Path,
    *,
    target_width: int | None,
    target_height: int | None,
    crf: int,
    preset: str,
    vcodec: str,
    acodec: str,
    abitrate: str,
    brightness: float,
    contrast: float,
    saturation: float,
    gamma: float,
    strip_metadata: bool,
    
    container: str | None,
 ) -> list[str]:

    color_filters = []
    if any([brightness != 0.0, contrast != 1.0, saturation != 1.0, gamma != 1.0]):
        color_filters.append(
            f"eq=brightness={brightness}:contrast={contrast}:saturation={saturation}:gamma={gamma}"
        )

    if target_width or target_height:
        w = target_width if target_width else -2
        h = target_height if target_height else -2
        if target_width and target_height:
            h = -2
        color_filters.append(f"scale={w}:{h}:flags=lanczos")

    vf_arg = ",".join(color_filters) if color_filters else None

    cmd = [ffmpeg_bin(), "-y", "-hide_banner", "-loglevel", "error", "-stats", "-i", str(inp)]

    # 1) Suppression de toutes les métadonnées originales + chapitres
    if strip_metadata:
        cmd += ["-map_metadata", "-1", "-map_chapters", "-1"]

    # 2) Mappage des streams principaux
    cmd += ["-map", "0:v?", "-map", "0:a?"]

    # 3) Paramètres d'encodage
    cmd += ["-c:v", vcodec, "-preset", preset, "-crf", str(crf)]
    if vf_arg:
        cmd += ["-vf", vf_arg]
    cmd += ["-c:a", acodec, "-b:a", abitrate]

    # 4) MOVFLAGS خاص بـ MP4/MOV (mdta + faststart)
    cont = (container or out.suffix.lower().lstrip(".")).lower()
    if cont in {"mp4", "mov", "m4v"}:
        # Only faststart; let ExifTool handle metadata atoms
        cmd += ["-movflags", "+faststart"]

    # 5) Sortie
    cmd.append(str(out))
    return cmd


def build_exiftool_cmd(
    out_path: Path,
    *,
    container: str | None,
    title: str | None,
    tags: list[str] | None,
    rating: str | None,
 ) -> list[str]:
    container_type = (container or out_path.suffix.lower().lstrip(".")).lower()
    tags_list = [t.strip() for t in (tags or []) if t and t.strip()]
    tags_joined = ", ".join(tags_list)

    cmd = [
        exiftool_bin(),
        "-m", "-overwrite_original",
        "-charset", "UTF8", "-charset", "filename=UTF8",
        "-sep", ", "
    ]

    if container_type in {"mp4", "mov", "m4v"}:
        # Nettoyage minimal compatible MP4
        cmd += [
            "-ItemList:Title=", "-QuickTime:Title=",
            "-Comment=", "-ItemList:Comment=",
            "-Keys:Keywords=", "-XMP-dc:Subject=",
            "-XMP-xmp:Rating=",
            "-Xtra:Title=", "-Xtra:Keywords=", "-Xtra:Rating="
        ]
    else:
        cmd += [
            "-Title=", "-Genre=", "-Comment=",
            "-XMP-dc:Subject=", "-XMP-xmp:Rating="
        ]
        if container_type in {"wmv", "wma", "asf"}:
            cmd += ["-RatingPercent=", "-XMP-microsoft:RatingPercent="]

    # Title
    if title:
        if container_type in {"mp4", "mov", "m4v"}:
            cmd += [
                f"-ItemList:Title={title}",
                f"-QuickTime:Title={title}",
                f"-XMP:Title={title}",
                f"-Xtra:Title={title}"
            ]
        else:
            cmd += [f"-Title={title}", f"-XMP:Title={title}"]

    # Tags
    if tags_list:
        if container_type in {"mp4", "mov", "m4v"}:
            cmd += [
                f"-Keys:Keywords={tags_joined}",
                f"-XMP-dc:Subject={tags_joined}",
                f"-Xtra:Keywords={tags_joined}"   # ← Explorer voit "Mots clés"
            ]
        else:
            cmd += [f"-Comment={tags_joined}"]
            for t in tags_list:
                cmd += [f"-XMP-dc:Subject={t}"]
            if container_type not in {"wmv","wma","asf"}:
                cmd += [f"-Genre={tags_joined}"]

    # Rating
    try:
        r = int(str(rating).strip()) if rating is not None else 5
    except Exception:
        r = 5

    if container_type in {"wmv","wma","asf"}:
        pct = {1: 1, 2: 25, 3: 50, 4: 75, 5: 99}.get(r, 99)
        cmd += [f"-RatingPercent={pct}", f"-XMP-microsoft:RatingPercent={pct}"]
    else:
        cmd += [f"-XMP-xmp:Rating={r}"]
        pct = {1: 1, 2: 25, 3: 50, 4: 75, 5: 99}.get(r, 99)
        cmd += [f"-Xtra:Rating={pct}"]      # ← Explorer voit "Notation"
        cmd += [f"-Keys:UserRating={r}"]    # ← Optionnel, compatibilité

    cmd.append(str(out_path))
    return cmd


#build exiftool cmd to set metadata to jpg
def build_exiftool_cmd_for_image(
    out_path: Path,
    *,
    title: str | None,
    tags: list[str] | None,
    rating: str | None
 ) -> list[str]:
    """
    JPEG/HEIC output (نستهدف JPG): نحيد كل الميتاداتا ونكتب Title/Keywords/Rating + Xtra باش يبان فـExplorer.
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
        "-all=",        # امسح كلشي
        "-P"            # حافظ على mtime إن أمكن
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

def check_metadata(cmd, log_print):
    log_print("[INFO] FFmpeg: aucune métadonnée écrite (ExifTool s'en charge).")
    return True

def dump_metadata_after_exiftool(path: Path, log_print, cont: str):
    try:
        if cont in {"mp4", "mov", "m4v"}:
            show = [
                exiftool_bin(), "-s", "-G1",
                "-ItemList:Title", "-QuickTime:Title", "-Xtra:Title",   # titres visibles Windows
                "-Keys:Keywords", "-XMP-dc:Subject", "-Xtra:Keywords",
                "-ItemList:Comment", "-QuickTime:Comment",
                "-Keys:UserRating", "-QuickTime:Rating", "-XMP-xmp:Rating", "-Xtra:Rating"
            ]
        elif cont in {"wmv", "wma", "asf"}:
            show = [
                exiftool_bin(), "-s", "-G1",
                "-Title", "-Genre", "-Comment",
                "-XMP-dc:Subject", "-ASF:RatingPercent"
            ]
        else:
            show = [
                exiftool_bin(), "-s", "-G1",
                "-Title", "-Genre", "-Comment",
                "-XMP-dc:Subject", "-XMP-xmp:Rating"
            ]
        show.append(str(path))
        res = subprocess.run(show, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        log_print("[META] " + res.stdout.strip())
    except Exception as e:
        log_print(f"[WARN] Impossible d'inspecter les métadonnées avec ExifTool: {e}")

def verify_written_metadata(path: Path, cont: str, title: str, tags: list[str], rating: str, log_print):
    try:
        show = [
            exiftool_bin(), "-j", "-G1",
            "-ItemList:Title", "-QuickTime:Title", "-Title", "-XMP-dc:Title",
            "-Keys:Keywords", "-XMP-dc:Subject", "-Comment",
            "-Keys:UserRating", "-XMP-xmp:Rating", "-ASF:RatingPercent",
            str(path)
        ]
        res = subprocess.run(show, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        import json
        arr = json.loads(res.stdout) if res.stdout.strip() else []
        if not arr:
            log_print("[VERIFY] Aucune métadonnée lue après écriture."); return
        d = arr[0]
        # titre
        actual_title = next((d.get(k) for k in ("ItemList:Title","QuickTime:Title","Title","XMP-dc:Title") if d.get(k)), "")
        # tags (Subject/Keywords)
        got = set()
        for k in ("XMP-dc:Subject","Keys:Keywords"):
            v = d.get(k)
            if isinstance(v, list): got |= {str(x).strip().lower() for x in v}
            elif isinstance(v, str): got |= {t.strip().lower() for t in v.split(",")}
        want = {t.strip().lower() for t in (tags or []) if t.strip()}
        # rating
        if cont in {"wmv","wma","asf"}:
            rating_ok = str(d.get("ASF:RatingPercent","")).strip() in {"99","100"}
        else:
            rating_ok = str(d.get("Keys:UserRating","")).strip() == str(rating) or \
                        str(d.get("XMP-xmp:Rating","")).strip() == str(rating)
        title_ok = (not title) or (str(actual_title).strip() == str(title).strip())
        tags_ok = (not want) or want.issubset(got)
        if title_ok and tags_ok and rating_ok:
            log_print(f"[VERIFY OK] Title='{actual_title}' | Tags={sorted(got)} | Rating={rating if cont not in {'wmv','wma','asf'} else '99%'}")
        else:
            log_print(f"[VERIFY FAIL] title_ok={title_ok} tags_ok={tags_ok} rating_ok={rating_ok}")
            log_print(f"  expected title='{title}', tags={sorted(want)}, rating={rating}")
            log_print(f"  got      title='{actual_title}', tags={sorted(got)}, rating={d.get('Keys:UserRating') or d.get('XMP-xmp:Rating') or d.get('ASF:RatingPercent')}")
    except Exception as e:
        log_print(f"[VERIFY ERR] {e}")

def process_one(src: Path, dst_root: Path, root: Path, args, log_print):
    rel = src.relative_to(root)

    csv_data = args.get("csv_data", {})
    folder_name = (rel.parts[0] if rel.parts else "").strip().lower()
    csv_row = csv_data.get(folder_name, {})

    csv_sku_kyopa = csv_row.get('sku kyopa', '').strip()
    csv_title = csv_row.get('title', '').strip()  # clés en minuscules après normalisation
    csv_tags = csv_row.get('tags', '').strip()

    title = csv_title or args["title"] or src.stem

    # def clean_filename(name):
    #     invalid_chars = '<>:"/\\|?*'
    #     for char in invalid_chars:
    #         name = name.replace(char, '_')
    #     return name or "video"

    def clean_filename(name):
        invalid = '<>:"/\\|?*'
        for ch in invalid:
            name = name.replace(ch, "_")   # ou " - "
        return (name or "video").rstrip(" .")[:150]

    safe_filename = clean_filename(title)

    if csv_sku_kyopa:
        out_folder = dst_root / csv_sku_kyopa
        out = out_folder / safe_filename
    else:
        out = (dst_root / rel.parent / safe_filename)

    if args["container"]:
        out = out.with_suffix("." + args["container"].lower())

    safe_mkdirs(out)

    if csv_tags:
        tags = [t.strip() for t in (csv_tags.split(",") if ',' in csv_tags else csv_tags.split()) if t.strip()]
        log_print(f"[INFO] Tags CSV pour {folder_name}: {tags}")
    else:
        tags = [t.strip() for t in args["tags"].split(",")] if args["tags"] else infer_tags_from_path(src, root)

    strip_metadata = True

    if not tags:
        log_print(f"[ATTENTION] Pas de tags trouvés pour {src.name}")
        tags = [folder_name]

    rating = "5"

    
    cmd = build_ffmpeg_cmd(
        src, out,
        target_width=args["width"],
        target_height=args["height"],
        crf=args["crf"],
        preset=args["preset"],
        vcodec=args["vcodec"],
        acodec=args["acodec"],
        abitrate=args["abitrate"],
        brightness=args["brightness"],
        contrast=args["contrast"],
        saturation=args["saturation"],
        gamma=args["gamma"],
        strip_metadata=strip_metadata,
        container=args["container"]
    )
    check_metadata(cmd, log_print)

    if out.exists() and not args["overwrite"]:
        log_print(f"[SKIP] Existe déjà: {out}")
        return

    log_print(f"[PROC] {src} -> {out} | Titre: '{title}' | Rating: {rating} | Tags: {tags}")
    try:
        subprocess.run(cmd, check=True)
        log_print(f"[OK] FFmpeg: {out}")

        cont = (args.get("container") or out.suffix.lower().lstrip(".")).lower()

        et_cmd = build_exiftool_cmd(
            out_path=out,
            container=cont,
            title=title,
            tags=tags,
            rating="5"
        )
        log_print(f"[INFO] ExifTool cmd: {' '.join(shlex.quote(c) for c in et_cmd)}")
        try:
            res = subprocess.run(et_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            log_print("[OK] ExifTool: " + (res.stdout.strip() or "métadonnées écrites."))
        except Exception as e:
            log_print(f"[WARN] ExifTool a échoué: {e}")
        finally:
            # Forcer affichage Explorer: écrire aussi via IPropertyStore
            if cont in {"mp4","m4v","mov"}:
                ok = set_win_explorer_props_mp4(str(out), title, tags, 5, log_print)
                if not ok:
                    log_print("[INFO] Windows properties non modifiables (lecture seule). Redémarrer Explorer peut aider.")
            try: os.utime(out, None)
            except Exception: pass

            # dump/verify (facultatif)
            try: dump_metadata_after_exiftool(out, log_print, cont)
            except Exception: pass
            try: verify_written_metadata(out, cont, title, tags, "5", log_print)
            except Exception: pass
    except subprocess.CalledProcessError as e:
        log_print(f"[ERROR] ffmpeg a échoué pour {src}:\n  {' '.join(shlex.quote(c) for c in cmd)}\n  {e}")
    except Exception as e:
        log_print(f"[ERROR] Erreur inattendue pour {src}: {e}")

def process_image_one(src: Path, dst_root: Path, root: Path, args, log_print):
    rel = src.relative_to(root)

    csv_data = args.get("csv_data", {})
    folder_name = (rel.parts[0] if rel.parts else "").strip().lower()
    csv_row = csv_data.get(folder_name, {}) or {}

    # Fallback: طابق بالاسم إذا ما لقاهاش بالفولدر
    if not csv_row:
        stem = src.stem.strip().lower()
        if stem in csv_data:
            csv_row = csv_data[stem]

    csv_sku_kyopa = (csv_row.get('sku kyopa', '') or "").strip()
    csv_title = (csv_row.get('title', '') or "").strip()
    csv_tags  = (csv_row.get('tags', '') or "").strip()

    title = csv_title or args.get("title") or src.stem

    def clean_filename(name):
        invalid = '<>:"/\\|?*'
        for ch in invalid:
            name = name.replace(ch, "_")
        return (name or "image").rstrip(" .")[:150]

    safe_filename = clean_filename(title)

    # نخرج JPG باش Explorer يبان فيه Rating/Keywords
    out_ext = ".jpg"

    # حدّد مجلد الخرج
    if csv_sku_kyopa:
        out_folder = dst_root / csv_sku_kyopa
    else:
        out_folder = dst_root / rel.parent

    # تأكد من وجود المجلد
    out_folder.mkdir(parents=True, exist_ok=True)

    # 🔢 دابا كنلقّاو الاسم المتاح بالتسلسل: title_1.jpg, title_2.jpg, ...
    i = 1
    while True:
        candidate = out_folder / f"{safe_filename}_{i}{out_ext}"
        if not candidate.exists():
            out = candidate
            break
        i += 1

    # tags
    if csv_tags:
        tags = [t.strip() for t in (csv_tags.split(",") if ',' in csv_tags else csv_tags.split()) if t.strip()]
        log_print(f"[INFO] Tags CSV pour {folder_name}: {tags}")
    else:
        tags = [t.strip() for t in (args.get("tags") or "").split(",")] if args.get("tags") else infer_tags_from_path(src, root)
    if not tags:
        tags = [folder_name] if folder_name else []

    rating = "5"

    log_print(f"[IMG] {src} -> {out} | Titre: '{title}' | Rating: {rating} | Tags: {tags}")

    try:
        # HEIC -> JPG، وإلا JPG أصلاً: ننسخو بالإسم الجديد
        if src.suffix.lower() == ".heic":
            convert_heic_to_jpg(src, out, quality=100)  # إذا بغيتي نقص للجودة دير 95
        else:
            shutil.copy2(src, out)

        # ExifTool: امسح metadata وكتب Title/Tags/Rating (+Xtra)
        et_cmd = build_exiftool_cmd_for_image(
            out_path=out,
            title=title,
            tags=tags,
            rating=rating
        )
        log_print(f"[INFO] ExifTool IMG cmd: {' '.join(shlex.quote(c) for c in et_cmd)}")
        try:
            res = subprocess.run(et_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            log_print("[OK] ExifTool IMG: " + (res.stdout.strip() or "métadonnées écrites."))
        except Exception as e:
            log_print(f"[WARN] ExifTool IMG a échoué: {e}")

        try:
            os.utime(out, None)
        except Exception:
            pass

    except Exception as e:
        log_print(f"[ERROR] Image: {e}")



# def run_batch(cfg, log_print, done_cb):
#     try:
#         in_root = Path(cfg["input_root"])
#         out_root = Path(cfg["output_root"])

#         if not in_root.exists():
#             log_print("❌ Le dossier d'entrée n'existe pas.")
#             return

#         files = [p for p in in_root.rglob("*") if p.suffix.lower() in VIDEO_EXTS]
#         if not files:
#             log_print("Aucune vidéo trouvée dans le dossier.")
#             return

#         # CSV
#         csv_data = {}
#         if "csv_path" in cfg and cfg["csv_path"]:
#             csv_data = read_csv_data(cfg["csv_path"])
#             if csv_data:
#                 log_print(f"Données CSV chargées: {len(csv_data)} entrées trouvées")
#             else:
#                 log_print("Aucune donnée CSV valide trouvée ou fichier non spécifié")
#         cfg["csv_data"] = csv_data

#         log_print(f"Traitement de {len(files)} vidéo(s)...")

#         for f in files:
#             if cfg["dry_run"]:
#                 rel = f.relative_to(in_root)
#                 folder_name = (rel.parts[0] if rel.parts else "").strip().lower()
#                 csv_row = csv_data.get(folder_name, {})
#                 csv_sku_kyopa = csv_row.get('sku kyopa', '').strip()
#                 csv_title = csv_row.get('title', '').strip()
#                 if csv_sku_kyopa:
#                     out = out_root / csv_sku_kyopa / rel.stem
#                 else:
#                     out = (out_root / rel.parent / rel.stem)
#                 if cfg["container"]:
#                     out = out.with_suffix("." + cfg["container"].lower())
#                 title = csv_title or cfg["title"] or f.stem
#                 if csv_row:
#                     log_print(f"[DRY] {f} -> {out} | title='{title}' | sku_kyopa='{csv_sku_kyopa}'")
#                 else:
#                     tags = [t.strip() for t in cfg["tags"].split(",")] if cfg["tags"] else infer_tags_from_path(f, in_root)
#                     log_print(f"[DRY] {f} -> {out} | title='{title}' | tags={tags}")
#             else:
#                 process_one(f, out_root, in_root, cfg, log_print)
#     finally:
#         done_cb()


# run batch forn images and videos 
def run_batch(cfg, log_print, done_cb):
    try:
        in_root = Path(cfg["input_root"])
        out_root = Path(cfg["output_root"])

        if not in_root.exists():
            log_print("❌ Le dossier d'entrée n'existe pas.")
            return

        all_files = [p for p in in_root.rglob("*") if p.is_file()]
        videos = [p for p in all_files if p.suffix.lower() in VIDEO_EXTS]
        images = [p for p in all_files if p.suffix.lower() in IMAGE_EXTS]

        if not videos and not images:
            log_print("Aucune vidéo ou image HEIC/JPG trouvée dans le dossier.")
            return

        # CSV
        csv_data = {}
        if "csv_path" in cfg and cfg["csv_path"]:
            csv_data = read_csv_data(cfg["csv_path"])
            if csv_data:
                log_print(f"Données CSV chargées: {len(csv_data)} entrées trouvées")
            else:
                log_print("Aucune donnée CSV valide trouvée ou fichier non spécifié")
        cfg["csv_data"] = csv_data

        # Filtrer les images selon le choix de l'utilisateur
        images_to_process = images if cfg.get("process_images", True) else []
        
        log_print(f"Traitement: {len(videos)} vidéo(s), {len(images_to_process)} image(s)...")

        if cfg["dry_run"]:
            for f in videos + images_to_process:
                rel = f.relative_to(in_root)
                folder_name = (rel.parts[0] if rel.parts else "").strip().lower()
                csv_row = csv_data.get(folder_name, {}) or {}
                if not csv_row:
                    stem = f.stem.strip().lower()
                    if stem in csv_data:
                        csv_row = csv_data[stem]
                csv_sku_kyopa = (csv_row.get('sku kyopa', '') or "").strip()
                csv_title = (csv_row.get('title', '') or "").strip()
                title = csv_title or cfg.get("title") or f.stem

                def clean_filename(name):
                    invalid = '<>:"/\\|?*'
                    for ch in invalid: name = name.replace(ch, "_")
                    return (name or "media").rstrip(" .")[:150]
                safe_filename = clean_filename(title)

                if f.suffix.lower() in VIDEO_EXTS:
                    out = (out_root / csv_sku_kyopa / safe_filename) if csv_sku_kyopa else (out_root / rel.parent / safe_filename)
                    if cfg["container"]:
                        out = out.with_suffix("." + cfg["container"].lower())
                else:
                    out_ext = ".jpg" if f.suffix.lower() == ".heic" else (f.suffix.lower() if f.suffix.lower() in {".jpg",".jpeg"} else ".jpg")
                    out = (out_root / csv_sku_kyopa / (safe_filename + out_ext)) if csv_sku_kyopa else (out_root / rel.parent / (safe_filename + out_ext))

                log_print(f"[DRY] {f} -> {out}")
            return

        # REAL RUN
        for f in videos:
            process_one(f, out_root, in_root, cfg, log_print)
        for f in images_to_process:
            process_image_one(f, out_root, in_root, cfg, log_print)

    finally:
        done_cb()




# =======================
#  Outil Fusion Dossiers
# =======================

def _copy_or_move_file(src: Path, dst: Path, mode: str, conflict: str) -> tuple[bool, Path]:
    """
    Copie ou déplace un fichier src vers dst en gérant les conflits.
    mode: "copy" | "move"
    conflict: "overwrite" | "skip" | "rename"
    Retourne (effectué, chemin_final)
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    target = dst
    if dst.exists():
        if conflict == "overwrite":
            pass
        elif conflict == "skip":
            return (False, dst)
        elif conflict == "rename":
            stem, suf = dst.stem, dst.suffix
            i = 2
            while target.exists():
                target = dst.with_name(f"{stem}_{i}{suf}")
                i += 1
        else:
            # par défaut, skip
            return (False, dst)

    if mode == "move":
        shutil.move(str(src), str(target))
    else:
        shutil.copy2(str(src), str(target))
    return (True, target)

def _merge_tree_one(src_dir: Path, dst_dir: Path, mode: str, conflict: str, log_print):
    for root, dirs, files in os.walk(src_dir):
        rel = Path(root).relative_to(src_dir)
        out_root = dst_dir / rel
        for d in dirs:
            (out_root / d).mkdir(parents=True, exist_ok=True)
        for f in files:
            s = Path(root) / f
            d = out_root / f
            done, final = _copy_or_move_file(s, d, mode, conflict)
            if done:
                log_print(f"[MERGE] {mode} {s} -> {final}")
            else:
                log_print(f"[MERGE] skip {s} (conflit)")

def merge_common_subdirs(parent_a: Path, parent_b: Path, dest_parent: Path, *, mode: str, conflict: str, log_print):
    """
    Trouve les sous-dossiers communs (même nom) dans parent_a et parent_b,
    puis fusionne leur contenu dans dest_parent/<nom>.
    mode: "copy" ou "move"
    conflict: "overwrite" | "skip" | "rename"
    """
    if not parent_a.exists() or not parent_b.exists():
        log_print("❌ Dossier parent introuvable."); return
    dest_parent.mkdir(parents=True, exist_ok=True)

    subs_a = {p.name for p in parent_a.iterdir() if p.is_dir()}
    subs_b = {p.name for p in parent_b.iterdir() if p.is_dir()}
    commons = sorted(subs_a & subs_b)
    if not commons:
        log_print("[MERGE] Aucun sous-dossier en commun."); return

    log_print(f"[MERGE] {len(commons)} sous-dossier(s) commun(s) trouvé(s): {commons}")
    for name in commons:
        src1 = parent_a / name
        src2 = parent_b / name
        dst = dest_parent / name
        dst.mkdir(parents=True, exist_ok=True)
        log_print(f"[MERGE] Fusion: '{name}'")
        _merge_tree_one(src1, dst, mode, conflict, log_print)
        _merge_tree_one(src2, dst, mode, conflict, log_print)
    log_print("[MERGE] Terminé.")

class MergeTool(tk.Toplevel):
    def __init__(self, master: tk.Tk):
        super().__init__(master)
        self.title("Fusion de dossiers (par nom)")
        self.geometry("780x520")
        self.resizable(True, True)

        # Toujours au-dessus et modale par rapport à la fenêtre principale
        self.transient(master)
        self.lift()
        self.attributes("-topmost", True)
        self.grab_set()
        self.focus_force()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Escape>", lambda e: self.on_close())
        self._center_on_parent()

        self.p1_var = tk.StringVar()
        self.p2_var = tk.StringVar()
        self.dst_var = tk.StringVar()

        frm = ttk.LabelFrame(self, text="Chemins")
        frm.pack(fill="x", padx=10, pady=8)

        self._row_path(frm, "Parent A:", self.p1_var, self.browse_p1)
        self._row_path(frm, "Parent B:", self.p2_var, self.browse_p2)
        self._row_path(frm, "Destination:", self.dst_var, self.browse_dst)

        opts = ttk.LabelFrame(self, text="Options")
        opts.pack(fill="x", padx=10, pady=8)

        self.mode_var = tk.StringVar(value="copy")
        self.conflict_var = tk.StringVar(value="rename")

        mrow = ttk.Frame(opts); mrow.pack(fill="x", padx=8, pady=4)
        ttk.Label(mrow, text="Mode:", width=14).pack(side="left")
        ttk.Radiobutton(mrow, text="Copier", value="copy", variable=self.mode_var).pack(side="left")
        ttk.Radiobutton(mrow, text="Déplacer", value="move", variable=self.mode_var).pack(side="left", padx=8)

        crow = ttk.Frame(opts); crow.pack(fill="x", padx=8, pady=4)
        ttk.Label(crow, text="Conflits:", width=14).pack(side="left")
        ttk.Radiobutton(crow, text="Renommer", value="rename", variable=self.conflict_var).pack(side="left")
        ttk.Radiobutton(crow, text="Écraser", value="overwrite", variable=self.conflict_var).pack(side="left", padx=8)
        ttk.Radiobutton(crow, text="Sauter", value="skip", variable=self.conflict_var).pack(side="left")

        runbar = ttk.Frame(self); runbar.pack(fill="x", padx=10, pady=6)
        self.start_btn = ttk.Button(runbar, text="Fusionner", command=self.start_merge)
        self.start_btn.pack(side="left")
        self.close_btn = ttk.Button(runbar, text="Fermer", command=self.destroy)
        self.close_btn.pack(side="left", padx=6)

        logf = ttk.LabelFrame(self, text="Journal")
        logf.pack(fill="both", expand=True, padx=10, pady=8)
        self.log = tk.Text(logf, height=16)
        self.log.pack(fill="both", expand=True)

    def _row_path(self, parent, label, var, cmd):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label, width=12).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Parcourir…", command=cmd).pack(side="left")

    def _append(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.update_idletasks()

    def _choose_dir(self, title: str) -> str | None:
        # S’assurer que la boîte de dialogue est au-dessus de la fenêtre Fusion
        try:
            self.attributes("-topmost", True)
            self.lift()
            self.focus_force()
            self.update_idletasks()
        except Exception:
            pass
        try:
            return filedialog.askdirectory(parent=self, title=title, mustexist=True)
        finally:
            try:
                self.lift()
                self.focus_force()
            except Exception:
                pass

    def browse_p1(self):
        d = self._choose_dir("Choisir le dossier parent A")
        if d: self.p1_var.set(d)

    def browse_p2(self):
        d = self._choose_dir("Choisir le dossier parent B")
        if d: self.p2_var.set(d)

    def browse_dst(self):
        d = self._choose_dir("Choisir le dossier de destination")
        if d: self.dst_var.set(d)

    def start_merge(self):
        p1 = self.p1_var.get().strip()
        p2 = self.p2_var.get().strip()
        dst = self.dst_var.get().strip()
        if not p1 or not p2 or not dst:
            messagebox.showerror("Erreur", "Sélectionnez Parent A, Parent B et Destination.")
            return
        mode = self.mode_var.get()
        conflict = self.conflict_var.get()
        self.start_btn.config(state="disabled")

        def log_print(msg): self._append(msg)

        def worker():
            try:
                merge_common_subdirs(Path(p1), Path(p2), Path(dst), mode=mode, conflict=conflict, log_print=log_print)
            finally:
                self.start_btn.config(state="normal")

        Thread(target=worker, daemon=True).start()

    def _center_on_parent(self):
        try:
            self.update_idletasks()
            pw = self.master.winfo_width()
            ph = self.master.winfo_height()
            px = self.master.winfo_rootx()
            py = self.master.winfo_rooty()
            w = self.winfo_width()
            h = self.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
            self.geometry(f"+{max(0,x)}+{max(0,y)}")
        except Exception:
            pass

    def on_close(self):
        try:
            self.grab_release()
        except Exception:
            pass
        try:
            self.attributes("-topmost", False)
        except Exception:
            pass
        # informer la fenêtre principale
        try:
            if hasattr(self.master, "merge_tool"):
                self.master.merge_tool = None
        except Exception:
            pass
        self.destroy()

# =======================
#  Tkinter GUI
# =======================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Traitement Vidéo par Lots")
        self.geometry("820x650")
        self.configure(bg="#2b2b2b")
        self._setup_modern_style()
        self.resizable(True, True)
        self.merge_tool = None

        # Paths row
        paths = ttk.LabelFrame(self, text="Chemins")
        paths.pack(fill="x", padx=10, pady=8)

        self.in_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.csv_var = tk.StringVar()

        self._row_path(paths, "Dossier d'entrée:", self.in_var, self.browse_in)
        self._row_path(paths, "Dossier de sortie:", self.out_var, self.browse_out)
        self._row_path(paths, "Fichier CSV:", self.csv_var, self.browse_csv)

        # Options
        opts = ttk.LabelFrame(self, text="Options")
        opts.pack(fill="x", padx=10, pady=8)

        # Left col
        left = ttk.Frame(opts)
        left.pack(side="left", fill="x", expand=True, padx=(0,10))

        self.width_var = tk.StringVar(value="2048")
        self.height_var = tk.StringVar(value="1536")
        self.crf_var = tk.IntVar(value=17)           # CRF
        self.preset_var = tk.StringVar(value="slow") # Preset
        self.vcodec_var = tk.StringVar(value="libx264")
        self.acodec_var = tk.StringVar(value="aac")
        self.abitrate_k_var = tk.IntVar(value=160)   # 320 kbps audio
        self.container_var = tk.StringVar(value="mp4")
        self.title_var = tk.StringVar()
        self.tags_var = tk.StringVar()

        self._row_entry(left, "Largeur:", self.width_var, "Ex: 1280")
        self._row_entry(left, "Hauteur:", self.height_var, "Souvent vide")
        self._row_scale(left, "CRF:", self.crf_var, 0, 51, resolution=1)
        # self._row_entry(left, "Preset:", self.preset_var, "ultrafast..placebo")
        # self._row_entry(left, "VCodec:", self.vcodec_var, "libx264/libx265/…")
        # self._row_entry(left, "Acodec:", self.acodec_var, "aac/…")
        # self._row_entry(left, "Débit audio:", self.abitrate_k_var, "320")

        # Right col
        right = ttk.Frame(opts)
        right.pack(side="left", fill="x", expand=True)

        self.brightness_var = tk.DoubleVar(value=0.0)
        self.contrast_var   = tk.DoubleVar(value=1.0)
        self.saturation_var = tk.DoubleVar(value=1.0)
        self.gamma_var      = tk.DoubleVar(value=1.0)

        self._row_scale(right, "Luminosité:", self.brightness_var, -1.0, 1.0, 0.01)
        self._row_scale(right, "Contraste:",   self.contrast_var,    0.0, 2.0,  0.01)
        self._row_scale(right, "Saturation:",  self.saturation_var,  0.0, 3.0,  0.01)
        self._row_scale(right, "Gamma:",       self.gamma_var,       0.1, 3.0,  0.01)
        self._row_scale(right, "Audio kbps:",  self.abitrate_k_var,  64,  160,  1)
        self._row_entry(right, "Conteneur:", self.container_var, "")
        # self._row_entry(right, "Titre:", self.title_var, "Optionnel")
        # self._row_entry(right, "Tags:", self.tags_var, "tag1,tag2")

        # Checkboxes
        toggles = ttk.Frame(self)
        toggles.pack(fill="x", padx=10, pady=4)
        self.keep_meta = tk.BooleanVar(value=False)
        self.overwrite = tk.BooleanVar(value=False)
        self.dry_run = tk.BooleanVar(value=False)
        self.process_images = tk.BooleanVar(value=True)

        ttk.Checkbutton(toggles, text="Traiter les images (HEIC/JPG)", variable=self.process_images).pack(side="left", padx=6)
        # ttk.Checkbutton(toggles, text="Conserver les métadonnées", variable=self.keep_meta).pack(side="left", padx=6)
        # ttk.Checkbutton(toggles, text="Écraser", variable=self.overwrite).pack(side="left", padx=6)
        # ttk.Checkbutton(toggles, text="Simulation", variable=self.dry_run).pack(side="left", padx=6)

        # Run buttons
        runbar = ttk.Frame(self)
        runbar.pack(fill="x", padx=10, pady=4)
        self.start_btn = ttk.Button(runbar, text="Démarrer", command=self.start_run, style='Success.TButton')
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(runbar, text="Arrêter (termine le fichier en cours)", command=self.request_stop, state="disabled", style='Danger.TButton')
        self.stop_btn.pack(side="left", padx=6)
        # Ajouter le bouton pour l’outil de fusion
        ttk.Button(runbar, text="Outil Fusion Dossiers…", command=self.open_merge_tool, style='Modern.TButton').pack(side="left", padx=6)
        attach_drive_csv_downloader(self, runbar_frame=runbar)
        # Log
        logf = ttk.LabelFrame(self, text="Journal")
        logf.pack(fill="both", expand=True, padx=10, pady=8)
        self.log = tk.Text(logf, height=18, bg="#1f1f1f", fg="#eaeaea", insertbackground="#eaeaea", borderwidth=0, highlightthickness=0)
        self.log.pack(fill="both", expand=True)
        self._append("Bienvenue ! Sélectionnez les dossiers d'entrée/sortie, le fichier CSV (optionnel), puis cliquez sur Démarrer.\n")

        self._stop_requested = False
        self._worker = None

    def _row_path(self, parent, label, var, cmd):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label, width=14).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Parcourir…", command=cmd).pack(side="left")

    def _row_entry(self, parent, label, var, placeholder=""):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=3)
        ttk.Label(row, text=label, width=14).pack(side="left")
        ent = ttk.Entry(row, textvariable=var)
        ent.pack(side="left", fill="x", expand=True)
        if placeholder:
            ent.insert(0, var.get())
        return ent

    def _row_scale(self, parent, label, var, from_, to_, resolution=0.01):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=3)
        ttk.Label(row, text=label, width=14).pack(side="left")
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

class BatchProcessorFrame(ttk.Frame):
    """Version intégrable de l'interface de traitement vidéo dans un conteneur Tk/ttk."""
    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        # Appliquer style moderne (reprend la config de App)
        self._setup_modern_style()
        # Conserver une ref pour outil de fusion
        self.merge_tool = None

        # Chemins
        paths = ttk.LabelFrame(self, text="Chemins")
        paths.pack(fill="x", padx=10, pady=8)

        self.in_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.csv_var = tk.StringVar()

        self._row_path(paths, "Dossier d'entrée:", self.in_var, self.browse_in)
        self._row_path(paths, "Dossier de sortie:", self.out_var, self.browse_out)
        self._row_path(paths, "Fichier CSV:", self.csv_var, self.browse_csv)

        # Options
        opts = ttk.LabelFrame(self, text="Options")
        opts.pack(fill="x", padx=10, pady=8)

        left = ttk.Frame(opts)
        left.pack(side="left", fill="x", expand=True, padx=(0,10))

        self.width_var = tk.StringVar()
        self.height_var = tk.StringVar()
        self.crf_var = tk.IntVar(value=17)           # CRF
        self.preset_var = tk.StringVar(value="slow") # Preset
        self.vcodec_var = tk.StringVar(value="libx264")
        self.acodec_var = tk.StringVar(value="aac")
        self.abitrate_k_var = tk.IntVar(value=160)   # 320 kbps audio
        self.container_var = tk.StringVar(value="mp4")
        self.title_var = tk.StringVar()
        self.tags_var = tk.StringVar()

        self._row_entry(left, "Largeur:", self.width_var, "Ex: 1280")
        self._row_entry(left, "Hauteur:", self.height_var, "Souvent vide")
        self._row_scale(left, "CRF:", self.crf_var, 0, 51, resolution=1)
        # self._row_entry(left, "Preset:", self.preset_var, "ultrafast..placebo")
        # self._row_entry(left, "VCodec:", self.vcodec_var, "libx264/libx265/…")
        # self._row_entry(left, "Acodec:", self.acodec_var, "aac/…")
        # self._row_entry(left, "Débit audio:", self.abitrate_k_var, "320")

        right = ttk.Frame(opts)
        right.pack(side="left", fill="x", expand=True)

        self.brightness_var = tk.DoubleVar(value=0.0)
        self.contrast_var   = tk.DoubleVar(value=1.0)
        self.saturation_var = tk.DoubleVar(value=1.0)
        self.gamma_var      = tk.DoubleVar(value=1.0)

        self._row_scale(right, "Luminosité:", self.brightness_var, -1.0, 1.0, 0.01)
        self._row_scale(right, "Contraste:",   self.contrast_var,    0.0, 2.0,  0.01)
        self._row_scale(right, "Saturation:",  self.saturation_var,  0.0, 3.0,  0.01)
        self._row_scale(right, "Gamma:",       self.gamma_var,       0.1, 3.0,  0.01)
        self._row_scale(right, "Audio kbps:",  self.abitrate_k_var,  64,  160,  1)
        self._row_entry(right, "Conteneur:", self.container_var, "")

        toggles = ttk.Frame(self)
        toggles.pack(fill="x", padx=10, pady=4)
        self.keep_meta = tk.BooleanVar(value=False)
        self.overwrite = tk.BooleanVar(value=False)
        self.dry_run = tk.BooleanVar(value=False)
        self.process_images = tk.BooleanVar(value=True)

        ttk.Checkbutton(toggles, text="Traiter les images (HEIC/JPG)", variable=self.process_images).pack(side="left", padx=6)

        runbar = ttk.Frame(self)
        runbar.pack(fill="x", padx=10, pady=4)
        self.start_btn = ttk.Button(runbar, text="Démarrer", command=self.start_run, style='Success.TButton')
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(runbar, text="Arrêter (termine le fichier en cours)", command=self.request_stop, state="disabled", style='Danger.TButton')
        self.stop_btn.pack(side="left", padx=6)
        ttk.Button(runbar, text="Outil Fusion Dossiers…", command=self.open_merge_tool, style='Modern.TButton').pack(side="left", padx=6)

        # Journal
        logf = ttk.LabelFrame(self, text="Journal")
        logf.pack(fill="both", expand=True, padx=10, pady=8)
        self.log = tk.Text(logf, height=18, bg="#1f1f1f", fg="#eaeaea", insertbackground="#eaeaea", borderwidth=0, highlightthickness=0)
        self.log.pack(fill="both", expand=True)
        self._append("Bienvenue ! Sélectionnez les dossiers d'entrée/sortie, le fichier CSV (optionnel), puis cliquez sur Démarrer.\n")

        self._stop_requested = False
        self._worker = None

    def _row_path(self, parent, label, var, cmd):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text=label, width=14).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Parcourir…", command=cmd).pack(side="left")

    def _row_entry(self, parent, label, var, placeholder=""):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=3)
        ttk.Label(row, text=label, width=14).pack(side="left")
        ent = ttk.Entry(row, textvariable=var)
        ent.pack(side="left", fill="x", expand=True)
        if placeholder:
            ent.insert(0, var.get())
        return ent

    def _row_scale(self, parent, label, var, from_, to_, resolution=0.01):
        row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=3)
        ttk.Label(row, text=label, width=14).pack(side="left")
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

    def _append(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.update_idletasks()

    def _setup_modern_style(self):
        try:
            style = ttk.Style()
            style.theme_use('clam')
            bg = '#2b2b2b'; fg = '#ffffff'; btn_bg = '#4a90e2'; btn_active = '#357abd'
            success_bg = '#28a745'; success_active = '#218838'; danger_bg = '#dc3545'; danger_active = '#c82333'
            style.configure('TFrame', background=bg)
            style.configure('TLabelframe', background=bg, foreground=fg)
            style.configure('TLabelframe.Label', background=bg, foreground=fg)
            style.configure('TLabel', background=bg, foreground=fg)
            style.configure('TButton', background=btn_bg, foreground=fg)
            style.map('TButton', background=[('active', btn_active)])
            style.configure('Modern.TButton', background=btn_bg, foreground=fg)
            style.map('Modern.TButton', background=[('active', btn_active)])
            style.configure('Success.TButton', background=success_bg, foreground=fg)
            style.map('Success.TButton', background=[('active', success_active)])
            style.configure('Danger.TButton', background=danger_bg, foreground=fg)
            style.map('Danger.TButton', background=[('active', danger_active)])
            style.configure('TEntry', fieldbackground='#3a3a3a', foreground=fg)
            style.configure('TNotebook', background=bg)
            style.configure('TNotebook.Tab', background='#3a3a3a', foreground=fg)
        except Exception:
            pass

    def browse_in(self):
        d = filedialog.askdirectory(title="Choisir le dossier d'entrée")
        if d:
            self.in_var.set(d)

    def browse_out(self):
        d = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if d:
            self.out_var.set(d)
            
    def browse_csv(self):
        d = filedialog.askopenfilename(title="Choisir le fichier CSV", 
                                      filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")])
        if d:
            self.csv_var.set(d)
            csv_data = read_csv_data(d)
            if csv_data:
                self._append(f"CSV chargé : {len(csv_data)} entrées trouvées.")
                self._append("Structure du CSV détectée :")
                sample = list(csv_data.items())[:2]
                for sku_orig, row in sample:
                    sku_kyopa = row.get('sku kyopa', '')
                    title = row.get('title', '')
                    tags = row.get('tags', '')
                    self._append(f"• '{sku_orig}' → '{sku_kyopa}'")
                    self._append(f"  - Titre: '{title}'")
                    self._append(f"  - Tags: '{tags}'")
            else:
                self._append("⚠️ Aucune entrée trouvée dans le CSV ou format incorrect.")

    def start_run(self):
        in_p = self.in_var.get().strip()
        out_p = self.out_var.get().strip()
        csv_p = self.csv_var.get().strip()
        if not in_p or not out_p:
            messagebox.showerror("Erreur", "Vous devez sélectionner les dossiers d'entrée et de sortie.")
            return

        def _to_int(s):
            s = s.strip(); return int(s) if s else None

        def _to_float(s, default):
            s = s.strip()
            try:
                return float(s) if s else default
            except:
                return default

        preset_val = self.preset_var.get().strip()
        vcodec_val = self.vcodec_var.get().strip()
        acodec_val = self.acodec_var.get().strip()
        if "medium" in preset_val and preset_val != "medium": preset_val = "medium"
        if "libx264" in vcodec_val and vcodec_val != "libx264": vcodec_val = "libx264"
        if "aac" in acodec_val and acodec_val != "aac": acodec_val = "aac"

        cfg = {
            "input_root": in_p,
            "output_root": out_p,
            "csv_path": csv_p,
            "width": _to_int(self.width_var.get()),
            "height": _to_int(self.height_var.get()),
            "crf": int(self.crf_var.get()),
            "preset": (self.preset_var.get() or "slow").strip(),
            "vcodec": (self.vcodec_var.get() or "libx264").strip(),
            "acodec": (self.acodec_var.get() or "aac").strip(),
            "abitrate": f"{int(self.abitrate_k_var.get())}k",
            "brightness": float(self.brightness_var.get()),
            "contrast":   float(self.contrast_var.get()),
            "saturation": float(self.saturation_var.get()),
            "gamma":      float(self.gamma_var.get()),
            "keep_metadata": False,
            "title": self.title_var.get().strip() or None,
            "tags": self.tags_var.get().strip() or None,
            "container": (self.container_var.get().strip() or "mp4").lower(),
            "overwrite": bool(self.overwrite.get()),
            "dry_run": bool(self.dry_run.get()),
            "process_images": bool(self.process_images.get()),
        }

        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._stop_requested = False
        self._append("Démarrage du traitement...")

        def log_print(msg):
            if self._stop_requested:
                pass
            self._append(msg)

        def done_cb():
            self._append("Terminé.")
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled")

        def worker():
            run_batch(cfg, log_print, done_cb)

        self._worker = Thread(target=worker, daemon=True)
        self._worker.start()

    def request_stop(self):
        self._stop_requested = True
        messagebox.showinfo("Info", "Le traitement s'arrêtera après avoir terminé le fichier en cours. Fermer la fenêtre annulera complètement.")

    def open_merge_tool(self):
        # Empêcher plusieurs copies; monter au premier plan
        top = self.winfo_toplevel()
        if self.merge_tool and self.merge_tool.winfo_exists():
            self.merge_tool.deiconify(); self.merge_tool.lift(); self.merge_tool.focus_force(); return
        self.merge_tool = MergeTool(top)
    def _setup_modern_style(self):
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
            style.configure('Modern.TButton', background=btn_bg, foreground=fg)
            style.map('Modern.TButton', background=[('active', btn_active)])
            style.configure('Success.TButton', background=success_bg, foreground=fg)
            style.map('Success.TButton', background=[('active', success_active)])
            style.configure('Danger.TButton', background=danger_bg, foreground=fg)
            style.map('Danger.TButton', background=[('active', danger_active)])
            # Entries
            style.configure('TEntry', fieldbackground='#3a3a3a', foreground=fg)
            # Notebook
            style.configure('TNotebook', background=bg)
            style.configure('TNotebook.Tab', background='#3a3a3a', foreground=fg)
        except Exception:
            pass

    def browse_in(self):
        d = filedialog.askdirectory(title="Choisir le dossier d'entrée")
        if d:
            self.in_var.set(d)

    def browse_out(self):
        d = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if d:
            self.out_var.set(d)
            
    def browse_csv(self):
        d = filedialog.askopenfilename(title="Choisir le fichier CSV", 
                                      filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")])
        if d:
            self.csv_var.set(d)
            csv_data = read_csv_data(d)
            if csv_data:
                self._append(f"CSV chargé : {len(csv_data)} entrées trouvées.")
                self._append("Structure du CSV détectée :")
                # adapter à la structure normalisée (clés en minuscules)
                sample = list(csv_data.items())[:2]
                for sku_orig, row in sample:
                    sku_kyopa = row.get('sku kyopa', '')
                    title = row.get('title', '')
                    tags = row.get('tags', '')
                    self._append(f"• '{sku_orig}' → '{sku_kyopa}'")
                    self._append(f"  - Titre: '{title}'")
                    self._append(f"  - Tags: '{tags}'")
            else:
                self._append("⚠️ Aucune entrée trouvée dans le CSV ou format incorrect.")

    def _append(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.update_idletasks()

    def start_run(self):
        in_p = self.in_var.get().strip()
        out_p = self.out_var.get().strip()
        csv_p = self.csv_var.get().strip()
        if not in_p or not out_p:
            messagebox.showerror("Erreur", "Vous devez sélectionner les dossiers d'entrée et de sortie.")
            return

        def _to_int(s):
            s = s.strip()
            return int(s) if s else None

        def _to_float(s, default):
            s = s.strip()
            try:
                return float(s) if s else default
            except:
                return default

        preset_val = self.preset_var.get().strip()
        vcodec_val = self.vcodec_var.get().strip()
        acodec_val = self.acodec_var.get().strip()
        if "medium" in preset_val and preset_val != "medium":
            preset_val = "medium"
        if "libx264" in vcodec_val and vcodec_val != "libx264":
            vcodec_val = "libx264"
        if "aac" in acodec_val and acodec_val != "aac":
            acodec_val = "aac"

        cfg = {
            "input_root": in_p,
            "output_root": out_p,
            "csv_path": csv_p,
            "width": _to_int(self.width_var.get()),
            "height": _to_int(self.height_var.get()),
            "crf": int(self.crf_var.get()),
            "preset": (self.preset_var.get() or "slow").strip(),
            "vcodec": (self.vcodec_var.get() or "libx264").strip(),
            "acodec": (self.acodec_var.get() or "aac").strip(),
            "abitrate": f"{int(self.abitrate_k_var.get())}k",
            "brightness": float(self.brightness_var.get()),
            "contrast":   float(self.contrast_var.get()),
            "saturation": float(self.saturation_var.get()),
            "gamma":      float(self.gamma_var.get()),
            "keep_metadata": False,
            "title": self.title_var.get().strip() or None,
            "tags": self.tags_var.get().strip() or None,
            "container": (self.container_var.get().strip() or "mp4").lower(),
            "overwrite": bool(self.overwrite.get()),
            "dry_run": bool(self.dry_run.get()),
            "process_images": bool(self.process_images.get()),
        }

        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._stop_requested = False
        self._append("Démarrage du traitement...")

        def log_print(msg):
            if self._stop_requested:
                pass
            self._append(msg)

        def done_cb():
            self._append("Terminé.")
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled")

        def worker():
            run_batch(cfg, log_print, done_cb)

        self._worker = Thread(target=worker, daemon=True)
        self._worker.start()

    def request_stop(self):
        self._stop_requested = True
        messagebox.showinfo("Info", "Le traitement s'arrêtera après avoir terminé le fichier en cours. Fermer la fenêtre annulera complètement.")

    def open_merge_tool(self):
        # empêcher plusieurs copies; garder toujours au-dessus de l’app
        if self.merge_tool and self.merge_tool.winfo_exists():
            self.merge_tool.deiconify()
            self.merge_tool.lift()
            self.merge_tool.focus_force()
            return
        self.merge_tool = MergeTool(self)

if __name__ == "__main__":
    App().mainloop()
