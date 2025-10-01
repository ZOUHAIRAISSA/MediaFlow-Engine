import os
from PIL import Image, ImageEnhance, ImageOps
import pillow_heif

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

# ---------- Main convert ----------
def convert_and_enhance(input_folder, output_folder, resize_width=None, preset="canva"):
    """
    Convert HEIC -> JPG, optionally enhance, keep metadata.
    """
    original_folder = os.path.join(output_folder, "Originals")
    enhanced_folder = os.path.join(output_folder, "Enhanced")
    os.makedirs(original_folder, exist_ok=True)
    os.makedirs(enhanced_folder, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if not file_name.lower().endswith(".heic"):
            continue

        input_path = os.path.join(input_folder, file_name)
        output_name = os.path.splitext(file_name)[0] + ".jpg"
        original_path = os.path.join(original_folder, output_name)
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

            # Save original-look
            img.save(original_path, **save_kwargs)
            print(f"Converted: {file_name} -> {output_name}")

            # Enhance
            if preset == "canva":
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

if __name__ == "__main__":
    input_folder = input("Enter the path to the folder containing HEIC files: ").strip()
    output_folder = input("Enter the path to the output folder: ").strip()
    rw = input("Enter the width to resize images (or press Enter to skip resizing): ").strip()
    resize_width = int(rw) if rw.isdigit() else None

    if not os.path.isdir(input_folder):
        print(f"Error: The folder {input_folder} does not exist or is not a directory.")
    else:
        convert_and_enhance(input_folder, output_folder, resize_width, preset="none")
        print("Conversion and enhancement process completed.")
