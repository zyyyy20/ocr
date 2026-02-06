import os
import sys
import traceback


IMAGE_PATH = r"C:\Users\Admin\Desktop\Snipaste_2026-02-06_10-50-16.png"


def _set_paddle_env_flags():
    os.environ.setdefault("FLAGS_use_mkldnn", "0")
    os.environ.setdefault("FLAGS_use_onednn", "0")
    os.environ.setdefault("FLAGS_enable_onednn", "0")
    os.environ.setdefault("FLAGS_disable_mkldnn", "1")
    os.environ.setdefault("FLAGS_enable_pir_in_executor", "0")
    os.environ.setdefault("FLAGS_use_new_executor", "0")


def _print_versions():
    print("Python:", sys.version.replace("\n", " "))
    try:
        import paddle

        print("paddle:", getattr(paddle, "__version__", "unknown"))
    except Exception as e:
        print("paddle: <import failed>", repr(e))

    try:
        import paddleocr

        print("paddleocr:", getattr(paddleocr, "__version__", "unknown"))
    except Exception as e:
        print("paddleocr: <import failed>", repr(e))


def _extract_lines(result):
    lines = []

    if not result:
        return lines

    if isinstance(result, list) and result and isinstance(result[0], list):
        for item in result[0]:
            try:
                text = item[1][0]
                score = float(item[1][1])
            except Exception:
                continue
            lines.append((text, score))
        return lines

    if isinstance(result, list):
        for item in result:
            if isinstance(item, dict):
                if "rec_texts" in item:
                    texts = item.get("rec_texts") or []
                    scores = item.get("rec_scores") or []
                    for i, t in enumerate(texts):
                        s = float(scores[i]) if i < len(scores) else None
                        lines.append((t, s))
                elif "text" in item:
                    lines.append((item.get("text"), item.get("score")))
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                try:
                    text = item[1][0]
                    score = float(item[1][1])
                    lines.append((text, score))
                except Exception:
                    pass

    return lines


def main():
    _set_paddle_env_flags()
    _print_versions()

    if not os.path.exists(IMAGE_PATH):
        print("错误：图片不存在：", IMAGE_PATH)
        sys.exit(2)
    print("image:", IMAGE_PATH)

    try:
        from paddleocr import PaddleOCR
    except Exception:
        traceback.print_exc()
        sys.exit(3)

    try:
        print("running_predict...")
        ocr = PaddleOCR(lang="ch")
        print("init_done")
        result = ocr.predict(
            IMAGE_PATH,
            use_textline_orientation=True,
        )
        print("predict_done")
    except TypeError:
        print("running_legacy_ocr...")
        ocr = PaddleOCR(use_angle_cls=True, lang="ch")
        result = ocr.ocr(IMAGE_PATH)
        print("legacy_ocr_done")
    except Exception as e:
        traceback.print_exc()
        msg = str(e)
        if "ConvertPirAttribute2RuntimeAttribute" in msg and "onednn_instruction" in msg:
            print(
                "\n检测到 Paddle 3.3.x 在 CPU(oneDNN) 推理的已知报错。"
                "\n可尝试降级 PaddlePaddle（例如 3.2.2）后再运行："
                "\n  pip install -U \"paddlepaddle==3.2.2\""
            )
        sys.exit(4)

    lines = _extract_lines(result)
    print("\n------------------ 识别结果 ------------------")
    if not lines:
        print("未识别到文字（或结果结构不匹配）。")
        print("raw_result_type:", type(result))
        print("raw_result_repr:", repr(result)[:2000])
        return

    for text, score in lines:
        if score is None:
            print(str(text))
        else:
            print(f"{text}\t{score:.4f}")


if __name__ == "__main__":
    main()
