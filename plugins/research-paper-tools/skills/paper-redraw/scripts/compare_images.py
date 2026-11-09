#!/usr/bin/env python3
"""Private QA comparison, never a redraw deliverable. Requires Pillow.

python compare_images.py reference.png result.png --output qa/pair.png
Optional: --reference-box x y w h --result-box x y w h --overlay qa/overlay.png
Aligns image height while preserving each aspect ratio. Does not compute a
scientific similarity score, alter inputs, or feed reference pixels to plotting.
"""

import argparse
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageOps


def read_image(path, box):
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source).convert("RGBA")
    if box:
        x, y, width, height = box
        if min(x, y) < 0 or min(width, height) <= 0 or x + width > image.width or y + height > image.height:
            raise ValueError(f"Crop exceeds image: {path}")
        image = image.crop((x, y, x + width, y + height))
    canvas = Image.new("RGBA", image.size, "white")
    return Image.alpha_composite(canvas, image).convert("RGB")


def compare(reference, result, output, overlay=None, height=900,
            reference_box=None, result_box=None, overwrite=False):
    reference, result, output = Path(reference), Path(result), Path(output)
    targets = [output] + ([Path(overlay)] if overlay else [])
    resolved = [p.resolve() for p in targets]
    if len(set(resolved)) != len(resolved) or any(p in (reference.resolve(), result.resolve()) for p in resolved):
        raise ValueError("QA output must not overwrite inputs or another output")
    if any(p.exists() for p in targets) and not overwrite:
        raise ValueError("QA output exists; use --overwrite for intended replacement")
    if not 50 <= height <= 2400:
        raise ValueError("height must be between 50 and 2400")
    images = [read_image(reference, reference_box), read_image(result, result_box)]
    sizes = [image.size for image in images]
    images = [image.resize((max(1, round(image.width * height / image.height)), height),
                           Image.Resampling.LANCZOS) for image in images]
    margin, header = 20, 30
    pair = Image.new("RGB", (sum(i.width for i in images) + 3 * margin, height + header + 2 * margin), "white")
    draw = ImageDraw.Draw(pair)
    x = margin
    for label, image in zip(("Reference (QA only)", "Rebuilt"), images):
        draw.text((x, margin), label, fill="black")
        pair.paste(image, (x, header + margin))
        x += image.width + margin
    output.parent.mkdir(parents=True, exist_ok=True)
    pair.save(output)
    if overlay:
        size = (max(image.width for image in images), height)
        canvases = []
        for image in images:
            canvas = Image.new("RGB", size, "white")
            canvas.paste(image, ((size[0] - image.width) // 2, 0))
            canvases.append(canvas)
        target = Path(overlay)
        target.parent.mkdir(parents=True, exist_ok=True)
        Image.blend(*canvases, 0.5).save(target)
    return {"comparison": str(output), "overlay": str(overlay) if overlay else None,
            "input_sizes": sizes, "aspect_preserved": True,
            "note": "QA reference pixels only; exclude these images from final deliverables."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("reference", type=Path)
    parser.add_argument("result", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overlay", type=Path)
    parser.add_argument("--height", type=int, default=900)
    parser.add_argument("--reference-box", type=int, nargs=4)
    parser.add_argument("--result-box", type=int, nargs=4)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    try:
        report = compare(args.reference, args.result, args.output, args.overlay, args.height,
                         args.reference_box, args.result_box, args.overwrite)
    except (OSError, ValueError) as error:
        parser.error(str(error))
    print(json.dumps(report))


if __name__ == "__main__":
    main()
