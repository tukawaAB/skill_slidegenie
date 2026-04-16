"""Image generation routing — dispatches to type-specific generators.

Ported from: ppt-addin/backend/services/make_pptx/image_gen/image_builder.py
"""
from PIL import Image

from slidegenie.gemini_client import GEMINIClient
from slidegenie.image_gen.graphic import GraphicImageGenerator
from slidegenie.image_gen.flow import FlowImageGenerator
from slidegenie.image_gen.matrix import MatrixImageGenerator


def prompt_to_image(
    make_type: str,
    prompt: str,
    logger,
    gemini_client: GEMINIClient,
    english_frag: bool = False,
    image: Image.Image | None = None,
    tone_manner: str | None = None,
) -> Image.Image:
    """Generate a slide image from prompt, dispatching to the appropriate generator."""

    if make_type == "graphic":
        gen = GraphicImageGenerator(
            gemini_client=gemini_client,
            logger=logger,
            prompt=prompt,
            english_frag=english_frag,
            image=image,
            tone_manner=tone_manner,
        )
    elif make_type == "flow":
        gen = FlowImageGenerator(
            gemini_client=gemini_client,
            logger=logger,
            prompt=prompt,
            english_frag=english_frag,
            image=image,
            tone_manner=tone_manner,
        )
    elif make_type == "matrix":
        gen = MatrixImageGenerator(
            gemini_client=gemini_client,
            logger=logger,
            prompt=prompt,
            english_frag=english_frag,
            image=image,
            tone_manner=tone_manner,
        )
    else:
        raise ValueError(f"Unsupported make_type: {make_type}. Use: graphic, flow, matrix")

    result = gen.make_image()
    if result is None:
        raise RuntimeError(f"Image generation returned None for make_type={make_type}")
    return result
