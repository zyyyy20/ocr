import os
import sys
import traceback

from PIL import Image, ImageDraw, ImageFont


IMAGE_PATH = r"C:\Users\Admin\Desktop\Snipaste_2026-02-06_16-15-11.png"
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "ocr_visualization.png")


def _set_paddle_env_flags():
    os.environ.setdefault("FLAGS_use_mkldnn", "0")
    os.environ.setdefault("FLAGS_use_onednn", "0")
    os.environ.setdefault("FLAGS_enable_onednn", "0")
    os.environ.setdefault("FLAGS_disable_mkldnn", "1")
    os.environ.setdefault("FLAGS_enable_pir_in_executor", "0")
    os.environ.setdefault("FLAGS_use_new_executor", "0")


def _get(obj, key, default=None):
    if obj is None:
        return default
    if isinstance(obj, dict):
        return obj.get(key, default)
    if hasattr(obj, "get"):
        try:
            v = obj.get(key)
            return default if v is None else v
        except Exception:
            pass
    if hasattr(obj, key):
        try:
            return getattr(obj, key)
        except Exception:
            pass
    try:
        return obj[key]
    except Exception:
        return default


def _as_boxes(raw):
    if raw is None:
        return []
    try:
        raw_list = list(raw)
    except Exception:
        return []

    boxes = []
    for item in raw_list:
        try:
            pts = list(item)
        except Exception:
            continue

        if len(pts) == 4:
            box = []
            ok = True
            for p in pts:
                try:
                    if hasattr(p, "__len__") and len(p) == 2:
                        x = float(p[0])
                        y = float(p[1])
                        box.append([x, y])
                    else:
                        ok = False
                        break
                except Exception:
                    ok = False
                    break
            if ok:
                boxes.append(box)
                continue

        if len(pts) == 4 and all(isinstance(v, (int, float)) for v in pts):
            x1, y1, x2, y2 = [float(v) for v in pts]
            boxes.append([[x1, y1], [x2, y1], [x2, y2], [x1, y2]])
            continue

        if len(pts) == 8 and all(isinstance(v, (int, float)) for v in pts):
            pts2 = [float(v) for v in pts]
            boxes.append([[pts2[0], pts2[1]], [pts2[2], pts2[3]], [pts2[4], pts2[5]], [pts2[6], pts2[7]]])
            continue

    return boxes


def _to_list(raw):
    if raw is None:
        return []
    try:
        return list(raw)
    except Exception:
        return []


def _find_font_path(ocr_item=None):
    vis_fonts = _get(ocr_item, "vis_fonts")
    if isinstance(vis_fonts, (list, tuple)):
        for p in vis_fonts:
            if isinstance(p, str) and os.path.exists(p):
                return p

    candidates = [
        os.path.join(os.path.dirname(__file__), "fonts", "simfang.ttf"),
        r"C:\Windows\Fonts\simfang.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def _draw_boxes_only(image, boxes):
    im = image.copy().convert("RGB")
    draw = ImageDraw.Draw(im)
    for box in boxes:
        if not box:
            continue
        poly = [(float(x), float(y)) for x, y in box]
        if len(poly) >= 2:
            draw.line(poly + [poly[0]], width=2, fill=(255, 0, 0))
    return im


def _draw_boxes_with_text(image, boxes, txts, scores, font_path=None):
    im = image.copy().convert("RGB")
    draw = ImageDraw.Draw(im)

    font_size = max(14, int(im.height * 0.02))
    font = None
    if font_path:
        try:
            font = ImageFont.truetype(font_path, font_size)
        except Exception:
            font = None
    if font is None:
        font = ImageFont.load_default()

    for i, box in enumerate(boxes):
        if not box:
            continue

        poly = [(float(x), float(y)) for x, y in box]
        if len(poly) >= 2:
            draw.line(poly + [poly[0]], width=2, fill=(255, 0, 0))

        if i >= len(txts):
            continue

        text = str(txts[i])
        if i < len(scores):
            try:
                text = f"{text} {float(scores[i]):.4f}"
            except Exception:
                pass

        x0 = min(p[0] for p in poly)
        y0 = min(p[1] for p in poly)

        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        pad = 2
        x1 = x0 + text_w + pad * 2
        y1 = y0 + text_h + pad * 2

        x0c = max(0, min(im.width - 1, int(x0)))
        y0c = max(0, min(im.height - 1, int(y0)))
        x1c = max(0, min(im.width, int(x1)))
        y1c = max(0, min(im.height, int(y1)))

        draw.rectangle([x0c, y0c, x1c, y1c], fill=(255, 255, 0))
        draw.text((x0c + pad, y0c + pad), text, fill=(0, 0, 0), font=font)

    return im


def main():
    _set_paddle_env_flags()

    if not os.path.exists(IMAGE_PATH):
        print("错误：图片不存在：", IMAGE_PATH)
        sys.exit(2)

    try:
        from paddleocr import PaddleOCR
    except Exception:
        traceback.print_exc()
        sys.exit(3)

    image = Image.open(IMAGE_PATH).convert("RGB")

    result = None
    ocr_item = None

    try:
        ocr = PaddleOCR(lang="ch")
        result = ocr.predict(IMAGE_PATH, use_textline_orientation=True)
        ocr_item = result[0] if result else None

        boxes = _as_boxes(_get(ocr_item, "dt_polys")) or _as_boxes(_get(ocr_item, "rec_polys")) or _as_boxes(
            _get(ocr_item, "rec_boxes")
        )
        txts = _to_list(_get(ocr_item, "rec_texts"))
        scores = _to_list(_get(ocr_item, "rec_scores"))
    except TypeError:
        ocr = PaddleOCR(use_angle_cls=True, lang="ch")
        legacy = ocr.ocr(IMAGE_PATH)
        legacy_lines = legacy[0] if legacy and legacy[0] else []
        boxes = [line[0] for line in legacy_lines]
        txts = [line[1][0] for line in legacy_lines]
        scores = [line[1][1] for line in legacy_lines]
    except Exception:
        traceback.print_exc()
        sys.exit(4)

    try:
        font_path = _find_font_path(ocr_item=ocr_item)
        print("boxes:", len(boxes), "txts:", len(txts), "scores:", len(scores), "font:", font_path)
        if boxes and txts:
            _draw_boxes_with_text(image, boxes, txts, scores, font_path).save(OUTPUT_PATH)
        else:
            _draw_boxes_only(image, boxes).save(OUTPUT_PATH)
    except Exception:
        traceback.print_exc()
        _draw_boxes_only(image, boxes).save(OUTPUT_PATH)

    print("已生成可视化图片：", OUTPUT_PATH)


if __name__ == "__main__":
    main()

