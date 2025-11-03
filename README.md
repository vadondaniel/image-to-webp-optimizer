# Image to WebP Optimizer

I like to archive the comics and manga I read, but over time I realized that the storage space they take up adds up quickly. I didn't want to lose visual quality just to save space, but since the chapters are usually stored as `.cbz` or `.zip` files, optimizing them manually was inconvenient.

To solve this, I created this program. After extracting a batch of chapters, it can automatically convert the images to **WebP** (a compact image format) and optionally reduce the image quality slightly (mostly by lowering the color depth or softening edges) to further save space. Once the images are processed, the program can **repack** each folder into an archive automatically, so there's no need to zip them one by one.

This process can typically reduce storage usage by **around 30-50%** with *very minimal visible quality loss*, and by **up to about 90%** with *slight quality degradation* (but without reducing resolution). Actual results vary depending on the source images.

Although it was designed with comic and manga archives in mind, it can also be used for more general image batch processing and compression tasks.

## Features

- Batch convert a single folder at a time, or enable Library Mode to handle each subfolder individually and output `.webp`.
- Adjustable quality slider that feeds directly into [`cwebp`](https://developers.google.com/speed/webp/download) so you can trade color fidelity for size where it makes sense.
- Optional skip for files that are already WebP to avoid double compression.
- Choice between keeping optimized copies alongside originals or replacing the originals in-place.
- Automatic re-packing into `.cbz` (comic book archive) or `.zip` once conversion finishes.
- Progress bar, cancel button, run summaries, and a 20-run history to replay past settings.

## Requirements

- Python 3.10+ (tested with CPython).
- Dependencies listed in `requirements.txt` (currently PyQt6 for the GUI).
- A [`cwebp`](https://developers.google.com/speed/webp/download) encoder in your `PATH` or in the project directory. A Windows build (`cwebp.exe`) is bundled here; on macOS/Linux install `libwebp` from your package manager.

## Getting Started

```powershell
git clone https://github.com/vadondaniel/image-to-webp-optimizer.git
cd "image-to-webp-optimizer"
# (Optional) python -m venv .venv; .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python image_optimizer.pyw
```

If you prefer WSL/Linux/macOS shells the commands are the same, just activate the virtual environment accordingly.

## Typical Manga/Comic Workflow

1. **Unzip the chapters you want to archive.** Each chapter should sit in its own folder (e.g., `Series/Chapter 12/` with JPG/PNG files inside).
2. **Launch the app** with `python image_optimizer.pyw` (or a double click) and click *Browse* to pick the parent folder.
3. Enable **Library Mode** when the parent folder contains one subfolder per chapter; the app will process each subfolder separately.
4. Set the **quality slider**. Values between 65–80 usually give me about 30-50% savings while keeping line art sharp. Higher values retain more color detail; 100 triggers lossless conversion for PNG sources.
5. Decide whether to **replace originals** (deletes the source JPG/PNG after the WebP copy is made) or keep them and let the app build new archives.
6. Pick **CBZ vs ZIP** packaging. CBZ works seamlessly with most comic readers; ZIP keeps things generic.
7. Optionally tick **Skip existing WebP files** to leave previously optimized pages untouched.
8. Hit **Start Conversion**. Watch progress per image, cancel any time, and review the run summary dialog for savings per folder.

When you keep archives, the tool writes them next to the original folders (e.g., `Chapter 12.cbz`) and cleans up the temporary working directory. If you replace originals, the WebP pages land back inside the original folder.

## General Usage Notes

- You can repurpose the tool for photo albums, documentation scans, or any folder that holds PNG/JPEG images directly (the app processes one folder level at a time).
- Run histories live in `image_optimizer.history.json`. Double-click entries in the UI to reuse the same folders and settings.
- Windows users can drop `image_optimizer.pyw` onto a Python launcher or pin a shortcut; the `.pyw` extension keeps the console hidden.
- If you see "`'cwebp' executable not found`", install `libwebp` or copy the [`cwebp`](https://developers.google.com/speed/webp/download) encoder binary next to `image_optimizer.pyw`.
