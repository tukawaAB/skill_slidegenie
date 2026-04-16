---
name: slidegenie
description: >
  Skill for working with the SlideGenie project — a local AI-powered PowerPoint slide generator
  that uses Gemini API to create professional consulting-style slides from text prompts.
  Use this skill whenever the user wants to: generate slides with slidegenie, run slidegenie CLI commands,
  modify or extend the slidegenie codebase (add slide types, edit prompts, change styles),
  debug slidegenie pipeline issues, or understand how the slidegenie architecture works.
  Also trigger when the user mentions: pptx generation, slide generation pipeline, Gemini OCR for slides,
  image-to-pptx conversion, or slide/deck generation from prompts.
---

# SlideGenie — AI Slide Generation Tool

SlideGenie is a Python CLI tool that generates professional PowerPoint slides from text prompts using Gemini API. It produces consulting-firm quality slides with a 4-step pipeline: language detection, slide type selection, image generation, and OCR-to-editable-PPTX conversion.

## Project Location (Self-Contained Skill)

This skill is **self-contained** — the entire project code ships inside the skill folder itself.

```
<SKILL_DIR>/                    # e.g. ~/.claude/skills/slidegenie/
├── SKILL.md                    # This file
├── README_SETUP.md             # First-time setup instructions (read this!)
├── .env.example                # Template for API key — copy to .env and fill in
├── pyproject.toml              # Python project manifest (uv)
├── uv.lock                     # Dependency lock file
├── discussion_slides.json      # Sample batch config
└── src/slidegenie/             # Source code (see file map below)
```

`<SKILL_DIR>` on a typical installation resolves to:
- Windows: `C:\Users\<user>\.claude\skills\slidegenie\`
- macOS/Linux: `~/.claude/skills/slidegenie/`

**Before first use**, follow `README_SETUP.md` to install dependencies and configure the Google API key.

## Quick Start

All commands are run from the skill directory (where `pyproject.toml` lives).

```bash
# Activate the virtual environment (created during setup)
cd <SKILL_DIR>
# Windows (Git Bash):   source .venv/Scripts/activate
# macOS/Linux:          source .venv/bin/activate

# API keys are loaded from .env file automatically (see .env.example)
# Or set directly: export GOOGLE_API_KEY="your-key"
# Or use ADC: gcloud auth application-default login

# Generate a single slide
python -m slidegenie generate -p "DX推進の3つのステップ" -o output.pptx

# Generate with specific type
python -m slidegenie generate -p "業務フロー改善" -o flow.pptx -t flow

# Image-only mode (no OCR)
python -m slidegenie generate -p "組織体制図" -o org.pptx -m image

# Convert existing image to editable PPTX
python -m slidegenie convert -i slide.png -o converted.pptx

# Batch generation from JSON config
python -m slidegenie batch -c specs.json -o deck.pptx -w 4

# Merge multiple PPTX files
python -m slidegenie merge -i file1.pptx -i file2.pptx -o merged.pptx
```

Batch config JSON format:
```json
[
  {"prompt": "DX推進の概要", "make_type": "graphic"},
  {"prompt": "導入プロセス", "make_type": "flow"},
  {"prompt": "競合比較", "make_type": "matrix"}
]
```

## Architecture Overview

The pipeline flows like this:

```
User Prompt
    │
    ├─ 1. Language Detection (is_english → en/ja)
    ├─ 2. Type Selection (Gemini LLM → graphic/flow/matrix)
    ├─ 3. Image Generation (Gemini gemini-3-pro-image-preview → PIL Image)
    └─ 4a. OCR (Gemini gemini-3-flash-preview → JSON with texts/icons/shapes)
       4b. Post-processing (font normalization, text alignment)
       4c. PPTX Build (python-pptx with template → .pptx file)
```

### Key Source Files

All paths are relative to `<SKILL_DIR>/`.

| File | Purpose |
|------|---------|
| `src/slidegenie/pipeline.py` | Main orchestrator — `generate_slide()`, `batch_generate()`, `convert_image()`, `merge_pptx()` |
| `src/slidegenie/cli.py` | Click CLI commands: generate, convert, batch, merge |
| `src/slidegenie/auth.py` | Gemini auth (GOOGLE_API_KEY or ADC via VertexAI) |
| `src/slidegenie/gemini_client.py` | `GEMINIClient` — image gen, OCR, chat completion with retry (10 retries, exponential backoff) |
| `src/slidegenie/image_gen/builder.py` | Routes to type-specific generators |
| `src/slidegenie/image_gen/graphic.py` | `GraphicImageGenerator` — diagrams, org charts |
| `src/slidegenie/image_gen/flow.py` | `FlowImageGenerator` — process flows, timelines |
| `src/slidegenie/image_gen/matrix.py` | `MatrixImageGenerator` — comparison tables, evaluation grids |
| `src/slidegenie/json_gen/builder.py` | `image_to_json()` — OCR image to structured JSON |
| `src/slidegenie/json_gen/postprocess.py` | Font size normalization, text alignment auto-detection |
| `src/slidegenie/slide_gen/builder.py` | `json_to_pptx()`, `multi_json_to_pptx()` — JSON to PPTX with template |
| `src/slidegenie/shapes/object_function.py` | Shape rendering, text fitting/truncation, validation |
| `src/slidegenie/utils/constants.py` | `PowerPointConfig`, `GEMINIConfig` |
| `src/slidegenie/utils/common.py` | `load_prompt()`, `is_english()`, `normalize_to_fullhd()` |
| `src/slidegenie/templates/template.pptx` | Base PPTX template with title/lead placeholders |

### Prompt Templates (Jinja2)

Located in `src/slidegenie/prompts/`:

| Template | Used For |
|----------|----------|
| `addin/select_maketype.j2` | LLM picks graphic/flow/matrix from user prompt |
| `addin/make_graphicimage_ja.j2` / `_en.j2` | Image gen prompt for graphic slides |
| `addin/make_flowimage_ja.j2` / `_en.j2` | Image gen prompt for flow slides |
| `addin/make_matriximage_ja.j2` / `_en.j2` | Image gen prompt for matrix slides |
| `addin/gemini_ocr_all.j2` | Unified OCR prompt (text + icon + shape detection) |
| `addin/build_json_only_text.j2` | Fallback: generate title/lead/body from prompt only |
| `tone-manner/_tone_and_manner_common_ja.j2` / `_en.j2` | Consulting-firm visual style guidelines |

Templates are rendered via `load_prompt(template_name, **kwargs)` which searches both `prompts/tone-manner/` and `prompts/addin/` directories.

## Design Details

### Slide Dimensions & Style
- **Size**: 13.3333 x 7.5 inches (16:9, standard widescreen)
- **Font**: Yu Gothic Light
- **Colors**: Taupe (#6D6358), Light Beige (#D9D0C1), White background, Dark Gray (#333333) text
- **Style**: Consulting-firm — minimal, clean, no gradients/shadows/3D, sharp rectangles

### Gemini Models
- **Image generation**: `gemini-3-pro-image-preview` (16:9, 1K resolution)
- **OCR / Chat**: `gemini-3-flash-preview`
- **Retry**: Up to 10 retries with exponential backoff (1s base, max 60s)

### OCR Coordinate System
- `box_2d` coordinates are normalized 0-1000, then unscaled to actual pixel dimensions
- Items classified as: `text` (with tag: Title/Lead/Body), `icon`, `shape`
- Shape types: RECTANGLE, ROUNDED_RECTANGLE, OVAL, DIAMOND, CHEVRON, PENTAGON, arrows, etc.
- Icons are cropped from the source image and embedded as pictures in the PPTX

### Text Fitting
- Auto font-size reduction: tries from initial size (default 18pt) down to 4pt
- If text still doesn't fit at 4pt, truncates with "..." ellipsis
- CJK/Latin aware: full-width chars get 1.2x weight, half-width 0.6x, spaces 0.3x
- Shape-specific text area ratios: OVAL 75%, PENTAGON 65%, RECTANGLE 100%
- 0.05 inch padding on all sides of text within shapes

### Batch Processing
- Phase 1: Parallel image generation + OCR (ThreadPoolExecutor, configurable workers)
- Phase 2: Sequential multi-slide PPTX assembly
- Template slide 0 reused; slides 1+ get cloned layout with title/lead textboxes

## Common Development Tasks

### Adding a New Slide Type

1. Create `src/slidegenie/image_gen/newtype.py` following the pattern of `graphic.py`:
   - Class with `__init__` taking gemini_client, logger, prompt, english_frag, image, tone_manner
   - `build_image_prompt()` method loading a Jinja2 template
   - `make_image()` method calling `gemini_client.generate_image()`

2. Create prompt templates:
   - `src/slidegenie/prompts/addin/make_newtypeimage_ja.j2`
   - `src/slidegenie/prompts/addin/make_newtypeimage_en.j2`

3. Register in `src/slidegenie/image_gen/builder.py` — add elif branch in `prompt_to_image()`

4. Update `src/slidegenie/prompts/addin/select_maketype.j2` to include the new type description

5. Add to valid types in `src/slidegenie/slide_gen/builder.py` `json_to_pptx()` make_type check

### Modifying Visual Style

Edit `src/slidegenie/prompts/tone-manner/_tone_and_manner_common_ja.j2` (or `_en.j2`):
- `color_palette` section for colors
- `typography` section for font styles and sizes
- `layout_geometry` for header/lead/margin ratios
- `graphics` section for icon/chart styling rules

### Adjusting Text Fitting Behavior

Edit `src/slidegenie/utils/constants.py` `PowerPointConfig`:
- `MIN_FONT_SIZE` / `MAX_FONT_SIZE` — font size range for auto-fitting
- `TEXT_PADDING_INCHES` — padding inside shapes
- `SHAPE_TEXT_AREA_RATIOS` — usable text area per shape type
- Character width ratios: `FULL_WIDTH_CHAR_WIDTH_RATIO`, `HALF_WIDTH_CHAR_WIDTH_RATIO`, etc.

### Debugging the Pipeline

Common issues and where to look:

| Issue | Check |
|-------|-------|
| Auth errors | `auth.py` — verify GOOGLE_API_KEY or ADC credentials |
| Image gen fails | `gemini_client.py` `generate_image()` — check model name, retry logs |
| OCR returns empty | `gemini_client.py` `perform_ai_ocr()` — check OCR prompt, image size |
| Shapes missing in PPTX | `shapes/object_function.py` `validate_shape_data()` — check validation logs |
| Text truncated | `shapes/object_function.py` `calculate_optimal_text_and_shape()` — adjust font/size |
| Wrong slide type selected | `prompts/addin/select_maketype.j2` — refine type descriptions |
| Coordinates off | `gemini_client.py` `box2d_unscale()` — check 0-1000 normalization |

Enable debug logging by setting logger level:
```python
from slidegenie.utils.logger import get_logger
logger = get_logger("pipeline")
logger.setLevel("DEBUG")
```
