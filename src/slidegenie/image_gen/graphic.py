"""Graphic (diagram) image generator.

Ported from: ppt-addin/backend/services/make_pptx/image_gen/image/make_graphic.py
"""
from PIL import Image

from slidegenie.gemini_client import GEMINIClient
from slidegenie.utils.constants import PowerPointConfig
from slidegenie.utils.common import load_prompt, resolve_image_prompt_template


class GraphicImageGenerator:
    def __init__(
        self,
        gemini_client: GEMINIClient,
        logger,
        prompt: str,
        english_frag: bool,
        image: Image.Image | None = None,
        tone_manner: str | None = None,
    ):
        self.gemini_client = gemini_client
        self.logger = logger
        self.prompt = prompt
        self.input_image = image
        self.english_frag = english_frag
        self.tone_manner = tone_manner
        self.image: Image.Image | None = None

    def build_image_prompt(self) -> str:
        template_name = resolve_image_prompt_template(
            english_frag=self.english_frag,
            addin_en_template="make_graphicimage_en",
            addin_ja_template="make_graphicimage_ja",
        )
        return load_prompt(
            template_name,
            prompt=self.prompt,
            default_font_name=PowerPointConfig.DEFAULT_FONT_NAME,
            tone_manner=self.tone_manner,
        )

    def make_image(self) -> Image.Image:
        prompt = self.build_image_prompt()
        self.image = self.gemini_client.generate_image(
            prompt=prompt,
            image=self.input_image,
        )
        return self.image
