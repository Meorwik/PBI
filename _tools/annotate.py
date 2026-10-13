"""Подписи на скриншотах: рамка + плашка с текстом (ImageMagick)."""
import subprocess


def annotate(src, dst, boxes, arrows=()):
    """boxes = [(color, (x1, y1, x2, y2), (tx, ty), text)], arrows = [(color, (tx, ty), text)] — только плашка."""
    args = ["magick", src, "-font", "DejaVu-Sans-Bold"]
    for color, (x1, y1, x2, y2), (tx, ty), text in boxes:
        args += ["-stroke", color, "-strokewidth", "5", "-fill", "none",
                 "-draw", f"rectangle {x1},{y1} {x2},{y2}"]
        args += _label(color, tx, ty, text)
    for color, (tx, ty), text in arrows:
        args += _label(color, tx, ty, text)
    subprocess.run(args + [dst], check=True)


def _label(color, tx, ty, text):
    return ["-stroke", "none", "-fill", "white", "-undercolor", color, "-pointsize", "28",
            "-annotate", f"+{tx}+{ty}", f" {text} "]
