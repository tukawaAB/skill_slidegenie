"""Gemini API client for image generation, OCR, and chat.

Ported from: ppt-addin/backend/services/azure_service.py (GEMINIClient class)
Stripped: Azure Blob/Table/DB dependencies, masking OCR mode (CV2 dependency)
Kept: unified OCR mode, retry logic, generate_image, generate_chat_completion
"""
import json
import re
import time
from io import BytesIO

from PIL import Image
from google import genai
from google.genai import types

from slidegenie.utils.constants import PowerPointConfig, GEMINIConfig
from slidegenie.utils.common import load_prompt
from slidegenie.utils.logger import get_logger


class GEMINIClient:
    def __init__(
        self,
        client: genai.Client,
        logger=None,
        ocr_model: str = None,
    ):
        self.client = client
        self.ocr_model = ocr_model or "gemini-3-flash-preview"
        self.icon_model = "gemini-3-flash-preview"
        self.chat_model = "gemini-3-flash-preview"
        self.image_model = "gemini-3-pro-image-preview"
        self.logger = logger or get_logger("gemini")

    def perform_ai_ocr(
        self,
        src_image: Image.Image,
        ocr_config: dict = None,
        unified_prompt: str = "gemini_ocr_all",
    ) -> dict:
        """Perform unified OCR: text + icon + shape detection in a single API call."""
        default_ocr_config = {
            "temperature": 0.01,
            "top_p": 0.88,
            "top_k": 25,
        }
        if ocr_config:
            default_ocr_config.update(ocr_config)

        def _normalize_box_2d(box):
            """Normalize box_2d to [y1, x1, y2, x2] int format."""
            if box is None:
                return None
            if isinstance(box, (list, tuple)) and len(box) == 1 and isinstance(box[0], (list, tuple)):
                box = box[0]
            if isinstance(box, str):
                nums = re.findall(r"-?\d+(?:\.\d+)?", box)
                if len(nums) < 4:
                    return None
                box = nums[:4]
            if not isinstance(box, (list, tuple)) or len(box) != 4:
                return None
            try:
                y1, x1, y2, x2 = [int(float(v)) for v in box]
            except (TypeError, ValueError):
                return None
            if y1 > y2:
                y1, y2 = y2, y1
            if x1 > x2:
                x1, x2 = x2, x1
            return [y1, x1, y2, x2]

        def _detect_item_type(data):
            raw_type = str(data.get("type", "")).strip().lower()
            if raw_type in {"text", "icon", "shape"}:
                return raw_type
            if data.get("shape_type"):
                return "shape"
            return ""

        def box2d_unscale(data, height, width):
            if isinstance(data, dict):
                if "box_2d" in data:
                    norm_box = _normalize_box_2d(data.get("box_2d"))
                    if norm_box is None:
                        data.pop("box_2d", None)
                    else:
                        ymin, xmin, ymax, xmax = norm_box
                        y1, x1 = int(ymin / 1000 * height), int(xmin / 1000 * width)
                        y2, x2 = int(ymax / 1000 * height), int(xmax / 1000 * width)
                        data["box_2d"] = [y1, x1, y2, x2]
                for key, value in data.items():
                    data[key] = box2d_unscale(value, height, width)
            elif isinstance(data, list):
                data = [box2d_unscale(item, height, width) for item in data]
            return data

        def box2d2xywh(data):
            if isinstance(data, dict):
                if "box_2d" in data:
                    norm_box = _normalize_box_2d(data.get("box_2d"))
                    if norm_box is None:
                        data.pop("box_2d", None)
                    else:
                        y1, x1, y2, x2 = norm_box
                        data["width"] = x2 - x1
                        data["height"] = y2 - y1
                        data["x"] = x1
                        data["y"] = y1
                        del data["box_2d"]
                for key, value in data.items():
                    data[key] = box2d2xywh(value)
            elif isinstance(data, list):
                data = [box2d2xywh(item) for item in data]
            return data

        def extract_icon_with_crop(original_image: Image.Image, icon_data):
            """Extract icons by cropping from the source image (PIL-based, no OpenCV)."""
            icons = []
            for item in icon_data:
                box = _normalize_box_2d(item.get("box_2d"))
                if box is None:
                    continue
                y1, x1, y2, x2 = box
                # Crop icon region from original image
                cropped = original_image.crop((x1, y1, x2, y2))
                if cropped.size[0] == 0 or cropped.size[1] == 0:
                    continue
                # Convert to RGBA and make near-background pixels transparent
                cropped = cropped.convert("RGBA")
                # Save to BytesIO
                icon_bytes = BytesIO()
                cropped.save(icon_bytes, format="PNG")
                icon_bytes.seek(0)
                icons.append({
                    "image_path": icon_bytes,
                    "x": x1,
                    "y": y1,
                    "width": x2 - x1,
                    "height": y2 - y1,
                    "shape_type": "AI_ICON",
                })
            return icons

        image = src_image
        if image.mode not in ("RGB",):
            image = image.convert("RGB")
        w, h = image.size

        # Resize for OCR if needed
        ocr_image = image.copy()
        if max(ocr_image.size) > PowerPointConfig.OCR_MAX_SIZE:
            ocr_image.thumbnail(
                (PowerPointConfig.OCR_MAX_SIZE, PowerPointConfig.OCR_MAX_SIZE),
                Image.Resampling.LANCZOS,
            )

        self.logger.info("perform AI-OCR (unified mode)")
        shape_types = "\n".join(PowerPointConfig.SHAPE_TYPE.values())
        prompt = load_prompt(unified_prompt, shape_types=shape_types)

        raw = self.generate_ocr_json(prompt=prompt, image=ocr_image, **default_ocr_config)
        if isinstance(raw, list):
            raw_list = raw
        elif isinstance(raw, dict):
            raw_list = raw.get("items", [])
            if not isinstance(raw_list, list):
                raw_list = next((v for v in raw.values() if isinstance(v, list)), [])
        else:
            raw_list = []

        all_data = box2d_unscale(raw_list, h, w)
        text_data = [d for d in all_data if _detect_item_type(d) == "text"]
        icon_data = [d for d in all_data if _detect_item_type(d) == "icon"]
        shape_data = [d for d in all_data if _detect_item_type(d) == "shape"]

        icons = extract_icon_with_crop(image, icon_data)
        return {
            "texts": box2d2xywh(text_data),
            "icons": icons,
            "shapes": box2d2xywh(shape_data),
        }

    def generate_image(
        self, prompt: str, image: Image.Image | None = None, **config_kwargs
    ) -> Image.Image | None:
        """Generate an image from prompt using Gemini (with retry logic)."""
        base_config = {
            "http_options": types.HttpOptions(timeout=400_000),
            "image_config": types.ImageConfig(
                aspect_ratio="16:9", image_size="1K"
            ),
            "response_modalities": ["IMAGE"],
        }
        if config_kwargs:
            base_config.update(config_kwargs)

        max_retries = GEMINIConfig.MAX_RETRIES
        base_delay = GEMINIConfig.BASE_DELAY

        for attempt in range(max_retries):
            try:
                response = self.client.models.generate_content(
                    model=self.image_model,
                    contents=[prompt, image] if image else [prompt],
                    config=types.GenerateContentConfig(**base_config),
                )

                for part in response.parts:
                    if part.inline_data is not None:
                        image_bytes = part.inline_data.data
                        pil_image = Image.open(BytesIO(image_bytes))
                        pil_image.load()
                        if attempt > 0:
                            self.logger.info(f"Image generation succeeded (retries: {attempt})")
                        return pil_image

                raise ValueError("No image was generated")

            except Exception as e:
                error_str = str(e).lower()
                non_retryable = ["authentication", "permission", "invalid",
                                 "validation", "badrequest", "unauthorized", "forbidden"]
                if any(err in error_str for err in non_retryable):
                    self.logger.error(f"Non-retryable error: {type(e).__name__}: {e}")
                    raise
                if attempt == max_retries - 1:
                    self.logger.error(f"Image generation failed {max_retries} times. Last error: {e}")
                    raise
                delay = min(base_delay * (2 ** attempt), 60.0)
                self.logger.warning(
                    f"Image generation failed (attempt {attempt + 1}/{max_retries}): {e}. "
                    f"Retrying in {delay:.1f}s..."
                )
                time.sleep(delay)

        raise Exception(f"Image generation failed {max_retries} times")

    def generate_chat_completion(
        self, prompt: str, image: Image.Image | None = None, **config_kwargs
    ) -> str:
        """Execute LLM call with retry logic."""
        base_config = {
            "http_options": types.HttpOptions(timeout=300_000),
            "top_p": 1.0,
            "temperature": 1.0,
            "response_modalities": ["TEXT"],
        }
        if config_kwargs:
            base_config.update(config_kwargs)

        max_retries = GEMINIConfig.MAX_RETRIES
        base_delay = GEMINIConfig.BASE_DELAY

        for attempt in range(max_retries):
            try:
                response = self.client.models.generate_content(
                    model=self.chat_model,
                    contents=[prompt, image] if image else [prompt],
                    config=types.GenerateContentConfig(**base_config),
                )
                if hasattr(response, "text"):
                    if attempt > 0:
                        self.logger.info(f"LLM call succeeded (retries: {attempt})")
                    return response.text
                if attempt > 0:
                    self.logger.info(f"LLM call succeeded (retries: {attempt})")
                return str(response)

            except Exception as e:
                error_str = str(e).lower()
                non_retryable = ["authentication", "permission", "invalid",
                                 "validation", "badrequest", "unauthorized", "forbidden"]
                if any(err in error_str for err in non_retryable):
                    self.logger.error(f"Non-retryable error: {type(e).__name__}: {e}")
                    raise
                if attempt == max_retries - 1:
                    self.logger.error(f"LLM call failed {max_retries} times. Last error: {e}")
                    raise
                delay = min(base_delay * (2 ** attempt), 60.0)
                self.logger.warning(
                    f"LLM call failed (attempt {attempt + 1}/{max_retries}): {e}. "
                    f"Retrying in {delay:.1f}s..."
                )
                time.sleep(delay)

        raise Exception(f"LLM call failed {max_retries} times")

    def generate_ocr_json(self, prompt: str, image: Image.Image | None = None, **config_kwargs):
        """Generate OCR JSON output with retry logic."""
        base_config = {
            "http_options": types.HttpOptions(timeout=300_000),
            "response_mime_type": "application/json",
        }
        if config_kwargs:
            base_config.update(config_kwargs)

        max_retries = GEMINIConfig.MAX_RETRIES
        base_delay = GEMINIConfig.BASE_DELAY

        for attempt in range(max_retries):
            try:
                response = self.client.models.generate_content(
                    model=self.ocr_model,
                    contents=[prompt, image] if image else [prompt],
                    config=types.GenerateContentConfig(**base_config),
                )
                if hasattr(response, "text") and response.text is not None:
                    result = json.loads(response.text)
                    if attempt > 0:
                        self.logger.info(f"OCR JSON generation succeeded (retries: {attempt})")
                    return result
                else:
                    raise ValueError(f"Response contains no text")

            except json.JSONDecodeError as e:
                if attempt == max_retries - 1:
                    self.logger.error(f"OCR JSON generation failed {max_retries} times (JSON parse error): {e}")
                    raise
                delay = min(base_delay * (2 ** attempt), 60.0)
                self.logger.warning(
                    f"OCR JSON parse error (attempt {attempt + 1}/{max_retries}): {e}. "
                    f"Retrying in {delay:.1f}s..."
                )
                time.sleep(delay)

            except Exception as e:
                error_str = str(e).lower()
                non_retryable = ["authentication", "permission", "invalid",
                                 "validation", "badrequest", "unauthorized", "forbidden"]
                if any(err in error_str for err in non_retryable):
                    self.logger.error(f"Non-retryable error: {type(e).__name__}: {e}")
                    raise
                if attempt == max_retries - 1:
                    self.logger.error(f"OCR JSON generation failed {max_retries} times. Last error: {e}")
                    raise
                delay = min(base_delay * (2 ** attempt), 60.0)
                self.logger.warning(
                    f"OCR JSON generation failed (attempt {attempt + 1}/{max_retries}): {e}. "
                    f"Retrying in {delay:.1f}s..."
                )
                time.sleep(delay)

        raise Exception(f"OCR JSON generation failed {max_retries} times")
