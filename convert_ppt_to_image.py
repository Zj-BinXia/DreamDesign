#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import importlib
import json
import math
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional, Union


def convert_ppt_to_images(
    ppt_path: Union[str, Path],
    out_dir: Union[str, Path],
    *,
    dpi: int = 160,
    image_format: str = "png",
    target_pixel_area: int = 2048 * 2048,
) -> List[Path]:
    """
    Convert a PPT/PPTX into one image per slide.

    Pipeline:
    1. Use LibreOffice/soffice headless to convert PPT/PPTX -> PDF
    2. Use PyMuPDF (fitz) to rasterize each PDF page -> image

    Requirements:
    - `soffice` or `libreoffice` available in PATH
    - `pip install pymupdf`
    """
    ppt_path = Path(ppt_path).expanduser().resolve()
    out_dir = Path(out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not ppt_path.exists():
        raise FileNotFoundError("PPT not found: {}".format(ppt_path))

    office_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not office_bin:
        raise RuntimeError("LibreOffice/soffice not found in PATH")

    fmt = str(image_format or "png").strip().lower()
    if fmt not in {"png", "jpg", "jpeg"}:
        raise ValueError("Unsupported image format: {}".format(image_format))

    try:
        fitz = importlib.import_module("fitz")
    except Exception as exc:
        raise RuntimeError("PyMuPDF not installed. Run: pip install pymupdf") from exc

    try:
        pil_image = importlib.import_module("PIL.Image")
    except Exception as exc:
        raise RuntimeError("Pillow not installed. Run: pip install pillow") from exc

    out_paths: List[Path] = []
    with tempfile.TemporaryDirectory(prefix="ppt2img_") as tmpdir:
        tmpdir_path = Path(tmpdir)
        cmd = [
            office_bin,
            "--headless",
            "--convert-to",
            "pdf:impress_pdf_Export",
            "--outdir",
            str(tmpdir_path),
            str(ppt_path),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0:
            raise RuntimeError(
                "LibreOffice conversion failed:\nstdout:\n{}\nstderr:\n{}".format(proc.stdout, proc.stderr)
            )

        pdf_path = tmpdir_path / "{}.pdf".format(ppt_path.stem)
        if not pdf_path.exists():
            raise RuntimeError("Converted PDF not found: {}".format(pdf_path))

        zoom = max(1, int(dpi)) / 72.0
        doc = fitz.open(str(pdf_path))
        try:
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                ext = "jpg" if fmt == "jpeg" else fmt
                out_path = out_dir / "{}_slide{:03d}.{}".format(ppt_path.stem, i, ext)
                pix.save(str(out_path))
                _resize_image_to_target_area(out_path, pil_image, target_pixel_area)
                out_paths.append(out_path)
        finally:
            doc.close()

    return out_paths


def convert_html_to_png(
    html_path: Union[str, Path],
    out_path: Union[str, Path],
    *,
    width: Optional[int] = None,
    height: Optional[int] = None,
    full_page: bool = True,
    scale: float = 1.0,
    selector: Optional[str] = ".slide",
    target_pixel_area: int = 2048 * 2048,
    supersample: float = 2.0,
) -> Path:
    """
    Render a local HTML file to PNG using Playwright.

    Requirements:
    - `pip install playwright`
    - Browser binary available to Playwright (run `playwright install chromium` if needed)
    """
    html_path = Path(html_path).expanduser().resolve()
    out_path = Path(out_path).expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if not html_path.exists():
        raise FileNotFoundError("HTML not found: {}".format(html_path))

    try:
        sync_api = importlib.import_module("playwright.sync_api")
    except Exception as exc:
        raise RuntimeError("Playwright not installed. Run: pip install playwright") from exc
    try:
        pil_image = importlib.import_module("PIL.Image")
    except Exception as exc:
        raise RuntimeError("Pillow not installed. Run: pip install pillow") from exc

    target_url = html_path.as_uri()
    viewport_w = max(1, int(width or 1280))
    viewport_h = max(1, int(height or 720))
    base_device_scale_factor = max(1.0, float(scale))

    # Avoid upscaling (which blurs). If the screenshot area is below target,
    # re-render with a higher device scale factor (supersampling), then only
    # downscale to match the target area.
    max_dsf = 8.0
    dsf = base_device_scale_factor
    tmp_out = out_path.with_suffix(".tmp.png")
    ss = max(1.0, float(supersample))
    target_render_area = int(round(float(target_pixel_area) * (ss * ss)))

    def _render_once(device_scale_factor: float) -> None:
        with sync_api.sync_playwright() as pw:
            browser = None
            last_error = None
            launch_attempts = [
                {"headless": True},
                {"headless": True, "channel": "chrome"},
                {"headless": True, "channel": "msedge"},
            ]
            for kwargs in launch_attempts:
                try:
                    browser = pw.chromium.launch(**kwargs)
                    break
                except Exception as exc:
                    last_error = exc
            if browser is None:
                raise RuntimeError(
                    "Unable to launch Playwright browser. Install one with: playwright install chromium"
                ) from last_error

            try:
                page = browser.new_page(
                    viewport={"width": viewport_w, "height": viewport_h},
                    device_scale_factor=max(1.0, float(device_scale_factor)),
                )
                page.goto(target_url, wait_until="networkidle")
                target = None
                if selector:
                    try:
                        locator = page.locator(selector).first
                        locator.wait_for(state="visible", timeout=5000)
                        target = locator
                    except Exception:
                        target = None
                if target is not None:
                    target.screenshot(path=str(tmp_out))
                else:
                    page.screenshot(path=str(tmp_out), full_page=bool(full_page))
            finally:
                browser.close()

    for _ in range(3):
        _render_once(dsf)
        with pil_image.open(str(tmp_out)) as im:
            w, h = im.size
            area = int(w) * int(h)
        if target_render_area > 0 and area < int(target_render_area * 0.98) and dsf < max_dsf:
            # Increase DSF to reach target area without upscaling.
            factor = math.sqrt(float(target_render_area) / float(max(1, area)))
            dsf = min(max_dsf, max(dsf * factor, dsf + 0.5))
            continue
        break

    tmp_out.replace(out_path)
    _resize_image_to_target_area(out_path, pil_image, target_pixel_area, allow_upscale=False)
    return out_path


def _resize_image_to_target_area(
    image_path: Union[str, Path],
    pil_image_module,
    target_pixel_area: int,
    *,
    allow_upscale: bool = True,
) -> Path:
    image_path = Path(image_path).expanduser().resolve()
    if target_pixel_area <= 0:
        return image_path
    with pil_image_module.open(str(image_path)) as im:
        w, h = im.size
        if w <= 0 or h <= 0:
            return image_path
        current_area = w * h
        if current_area <= 0:
            return image_path
        scale = math.sqrt(float(target_pixel_area) / float(current_area))
        if scale > 1.0 and not allow_upscale:
            return image_path
        new_w = max(1, int(round(w * scale)))
        new_h = max(1, int(round(h * scale)))
        if (new_w, new_h) == (w, h):
            return image_path
        resized = im.resize((new_w, new_h), resample=pil_image_module.LANCZOS)
        resized.save(str(image_path))
    return image_path


def main() -> None:
    ap = argparse.ArgumentParser(description="Convert PPT/PPTX or HTML into PNG images.")
    sub = ap.add_subparsers(dest="command", required=True)

    ap_ppt = sub.add_parser("ppt", help="Convert PPT/PPTX into one image per slide")
    ap_ppt.add_argument("ppt_path", help="Input .ppt or .pptx file")
    ap_ppt.add_argument("-o", "--out-dir", required=True, help="Output directory for rendered images")
    ap_ppt.add_argument("--dpi", type=int, default=160, help="Rasterization DPI (default: 160)")
    ap_ppt.add_argument("--format", default="png", choices=["png", "jpg", "jpeg"], help="Image format")
    ap_ppt.add_argument("--target-area", type=int, default=2048 * 2048, help="Target pixel area (default: 2048*2048)")

    ap_html = sub.add_parser("html", help="Render a local HTML file to PNG")
    ap_html.add_argument("html_path", help="Input .html file")
    ap_html.add_argument("-o", "--out-path", required=True, help="Output .png file path")
    ap_html.add_argument("--width", type=int, default=1280, help="Viewport width")
    ap_html.add_argument("--height", type=int, default=720, help="Viewport height")
    ap_html.add_argument("--scale", type=float, default=1.0, help="Device scale factor")
    ap_html.add_argument("--no-full-page", action="store_true", help="Only capture current viewport")
    ap_html.add_argument("--selector", default=".slide", help="CSS selector to capture exactly (default: .slide)")
    ap_html.add_argument("--target-area", type=int, default=2048 * 2048, help="Target pixel area (default: 2048*2048)")
    ap_html.add_argument("--supersample", type=float, default=2.0, help="Linear supersample factor before downscale (default: 2.0)")

    args = ap.parse_args()

    if args.command == "ppt":
        paths = convert_ppt_to_images(
            args.ppt_path,
            args.out_dir,
            dpi=int(args.dpi),
            image_format=str(args.format),
            target_pixel_area=int(args.target_area),
        )
        print(json.dumps([str(p) for p in paths], ensure_ascii=False, indent=2))
        return

    out_path = convert_html_to_png(
        args.html_path,
        args.out_path,
        width=int(args.width),
        height=int(args.height),
        full_page=not bool(args.no_full_page),
        scale=float(args.scale),
        selector=str(args.selector) if args.selector else None,
        target_pixel_area=int(args.target_area),
        supersample=float(args.supersample),
    )
    print(out_path)


if __name__ == "__main__":
    main()
