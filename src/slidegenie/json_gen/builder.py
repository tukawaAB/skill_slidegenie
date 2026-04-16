"""Image to JSON conversion via OCR.

Ported from: ppt-addin/backend/services/make_pptx/json_gen/json_builder.py
"""
import re
import json

from PIL import Image

from slidegenie.gemini_client import GEMINIClient
from slidegenie.utils.constants import PowerPointConfig
from slidegenie.utils.common import load_prompt


def image_to_json(
    make_type: str,
    logger,
    gemini_client: GEMINIClient,
    prompt: str | None = None,
    src_image: Image.Image | None = None,
) -> dict:
    """OCR an image and build a JSON structure for PPTX conversion."""

    def _hex_to_rgb(hex_str, default_none=False):
        if hex_str is None or hex_str == "":
            return None if default_none else (0, 0, 0)
        hex_str = str(hex_str).lstrip("#")
        if len(hex_str) != 6:
            return None if default_none else (0, 0, 0)
        try:
            return tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4))
        except (ValueError, IndexError):
            return None if default_none else (0, 0, 0)

    def _extract_json_object(text):
        if not text:
            return None
        fenced = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text, flags=re.IGNORECASE)
        candidate = fenced.group(1) if fenced else text
        start = candidate.find("{")
        end = candidate.rfind("}")
        if start == -1 or end == -1 or end <= start:
            return None
        try:
            return json.loads(candidate[start:end + 1])
        except Exception:
            return None

    def _generate_text_only_from_prompt(prompt, make_type, gemini_client):
        gen_prompt = load_prompt("build_json_only_text", prompt=prompt, make_type=make_type)
        raw = gemini_client.generate_chat_completion(prompt=gen_prompt)
        data = _extract_json_object(raw) or {}
        return (
            str(data.get("title") or "").strip(),
            str(data.get("lead") or "").strip(),
            str(data.get("body") or "").strip(),
        )

    try:
        ocr_data = gemini_client.perform_ai_ocr(
            src_image,
            unified_prompt="gemini_ocr_all",
        )

        logger.info(
            f"[json_builder] OCR result - texts: {len(ocr_data.get('texts', []))}, "
            f"icons: {len(ocr_data.get('icons', []))}, "
            f"shapes: {len(ocr_data.get('shapes', []))}"
        )

        cfg = PowerPointConfig.OBJECT_FUNCTION_CONFIG
        default_shape_type = cfg["DEFAULT_SHAPE_TYPE"]
        ai_icon_type = cfg["SHAPE_TYPE_AI_ICON"]

        json_data = {"title": None, "lead": None, "body": []}

        # Shapes
        json_data["body"].extend([
            {
                **shape,
                "shape_type": shape.get("shape_type") or default_shape_type,
                "fill_color": _hex_to_rgb(shape.get("fill_color"), default_none=True),
                "line_color": _hex_to_rgb(shape.get("line_color"), default_none=True),
                "font_color": _hex_to_rgb(shape.get("font_color")) if shape.get("font_color") else shape.get("font_color"),
                "border_width": 2.25 if shape.get("line_color") else 0,
            }
            for shape in ocr_data["shapes"]
        ])

        # Icons
        json_data["body"].extend([
            {**icon, "shape_type": ai_icon_type}
            for icon in ocr_data["icons"]
        ])

        # Texts
        for text_data in ocr_data["texts"]:
            tag = text_data.pop("tag", None)
            if tag == "Title":
                json_data["title"] = text_data.get("text", "")
            elif tag == "Lead":
                if json_data["lead"] is None:
                    json_data["lead"] = text_data.get("text", "")
                else:
                    json_data["body"].append({
                        **text_data,
                        "font_color": _hex_to_rgb(text_data.get("font_color")),
                        "shape_type": default_shape_type,
                    })
            else:
                json_data["body"].append({
                    **text_data,
                    "font_color": _hex_to_rgb(text_data.get("font_color")),
                    "shape_type": default_shape_type,
                })

        logger.info(
            f"[json_builder] Final: title={'set' if json_data['title'] else 'None'}, "
            f"lead={'set' if json_data['lead'] else 'None'}, "
            f"body={len(json_data['body'])} items"
        )
        return json_data

    except Exception as e:
        logger.error(f"[ERROR] Failed to generate image→json: {e}")

        title = lead = body_text = ""
        if prompt:
            try:
                title, lead, body_text = _generate_text_only_from_prompt(
                    prompt=prompt, make_type=make_type, gemini_client=gemini_client,
                )
            except Exception as llm_e:
                logger.warning(f"[WARN] Fallback prompt generation failed: {llm_e}")

        if not body_text:
            body_text = (prompt or "").strip()

        ppi_1k = PowerPointConfig.PPI_1K
        return {
            "_fallback": True,
            "title": title or "",
            "lead": lead or "",
            "body": [{
                "shape_type": PowerPointConfig.OBJECT_FUNCTION_CONFIG["SHAPE_TYPE_TEXTBOX"],
                "x": PowerPointConfig.FALLBACK_X * ppi_1k,
                "y": PowerPointConfig.FALLBACK_Y * ppi_1k,
                "width": PowerPointConfig.FALLBACK_WIDTH * ppi_1k,
                "height": PowerPointConfig.FALLBACK_HEIGHT * ppi_1k,
                "text": body_text or "",
                "font_size": PowerPointConfig.FALLBACK_FONT_SIZE,
                "font_color": list(PowerPointConfig.FALLBACK_FONT_COLOR),
                "text_align": PowerPointConfig.FALLBACK_TEXT_ALIGN,
                "vertical_align": PowerPointConfig.FALLBACK_VERTICAL_ALIGN,
            }],
        }
