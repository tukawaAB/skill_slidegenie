"""PowerPoint and Gemini configuration constants.

Ported from: ppt-addin/backend/constants/llm_constants.py
"""
import re


class PowerPointConfig:
    """Constants for PowerPoint file generation"""

    # Slide dimensions (16:9 aspect ratio)
    SLIDE_WIDTH_INCHES = 13.3333
    SLIDE_HEIGHT_INCHES = 7.5
    BLANK_SLIDE_LAYOUT_INDEX = 1

    # Font settings
    DEFAULT_FONT_SIZE = 18
    DEFAULT_FONT_NAME = "Yu Gothic Light"
    DEFAULT_FONT_COLOR = (0, 0, 0)
    DEFAULT_TEXT_ALIGNMENT = "center"
    DEFAULT_BORDER_WIDTH = 0

    # Font size limits
    MIN_FONT_SIZE = 4
    MAX_FONT_SIZE = 18
    FONT_SIZE_STEP = 1
    POINTS_PER_INCH = 72

    # Character width ratios (relative to font size)
    FULL_WIDTH_CHAR_WIDTH_RATIO = 1.2
    HALF_WIDTH_CHAR_WIDTH_RATIO = 0.6
    CJK_CHAR_WIDTH_RATIO = 0.8
    SPACE_CHAR_WIDTH_RATIO = 0.3
    CJK_DEFAULT_CHAR_WIDTH_RATIO = 1.0

    # Line height ratio
    LINE_HEIGHT_RATIO = 1

    # Text padding (in inches)
    TEXT_PADDING_INCHES = 0.05

    # Text truncation
    ELLIPSIS_TEXT = "..."
    ELLIPSIS_CHAR_COUNT = 3

    # Shape text area ratios (% of bounding box for text)
    SHAPE_TEXT_AREA_RATIOS = {
        "RECTANGLE": 1.0,
        "ROUNDED_RECTANGLE": 1.0,
        "OVAL": 0.75,
        "PENTAGON": 0.65,
        "DEFAULT": 0.8,
    }

    # SHAPE_TYPE
    SHAPE_TYPE = {
        "RECTANGLE": "RECTANGLE",
        "ROUNDED_RECTANGLE": "ROUNDED_RECTANGLE",
        "OVAL": "OVAL",
        "ISOSCELES_TRIANGLE": "ISOSCELES_TRIANGLE",
        "RIGHT_TRIANGLE": "RIGHT_TRIANGLE",
        "DIAMOND": "DIAMOND",
        "CHEVRON": "CHEVRON",
        "PENTAGON": "PENTAGON",
        "DOWN_ARROW": "DOWN_ARROW",
        "UP_ARROW": "UP_ARROW",
        "LEFT_ARROW": "LEFT_ARROW",
        "RIGHT_ARROW": "RIGHT_ARROW",
        "LEFT_RIGHT_ARROW": "LEFT_RIGHT_ARROW",
        "ARROW_PENTAGON": "ARROW_PENTAGON",
        "CYLINDER": "CYLINDER",
        "FLOWCHART_DOCUMENT": "FLOWCHART_DOCUMENT",
    }

    # Color labels
    COLOR_LABELS = {
        "ライトベージュ": (240, 235, 227),
        "黒":            (0,   0,   0),
        "白":            (255, 255, 255),
        "ダークグレー":   (51,  51,  51),
        "ミディアムグレー": (128, 128, 128),
        "ライトグレー":   (180, 180, 180),
        "ダークブルー":   (0,   25,  100),
        "ベージュ":       (200, 190, 170),
        "ダークベージュ": (125, 110, 90),
    }

    # RGB color validation
    RGB_MIN_VALUE = 0
    RGB_MAX_VALUE = 255
    RGB_COMPONENTS_COUNT = 3

    # Required shape properties
    REQUIRED_PROPERTIES = ["shape_type", "width", "height", "x", "y"]
    NUMERIC_POSITIVE_PROPERTIES = ["width", "height"]

    # Object function configuration
    OBJECT_FUNCTION_CONFIG = {
        "TEXT_PREVIEW_LENGTH": 30,
        "DEFAULT_SHAPE_TYPE": "RECTANGLE",
        "DEFAULT_ICON_NAME": "process.png",
        "FALLBACK_ICON_TEXT": "🔧",
        "MIN_CHARS_PER_LINE": 1,
        "MIN_LINES_NEEDED": 1,
        "PADDING_REDUCTION_FACTOR": 2,
        "MIN_TEXT_WIDTH": 0,
        "MIN_TEXT_HEIGHT": 0,
        "MIN_TEXT_LINES": 0,
        "MAX_TEXT_LINES": 1,
        "MIN_CHAR_WEIGHT": 0.0,
        "MIN_TRUNCATE_INDEX": 0,
        "SPACE_CHAR_WEIGHT_OFFSET": 1,
        "SHAPE_TYPE_TEXTBOX": "TEXTBOX",
        "SHAPE_TYPE_AI_ICON": "AI_ICON",
        "SHAPE_TYPE_ROUNDED_RECTANGLE": "ROUNDED_RECTANGLE",
        "ICON_BLOB_PATH_KEY": "ICON_BLOB_PATH",
        "AUTO_CALCULATION_DISABLED": True,
        "TEXT_FITS_DEFAULT": True,
        "RESIZED_DEFAULT": False,
    }

    # SlideBuilder
    EMU_PER_INCH = 914400
    PAD_LEFT_IN = 0.25
    PAD_RIGHT_IN = 0.25
    PAD_BOTTOM_IN = 0.25
    FOOTER_RESERVED = 0.60
    HEADER_BODY_GAP = 0.50
    CONTENT_TOP = 0.30
    FULLHD_W = 1920
    FULLHD_H = 1080

    # OCR Processing Constants
    OCR_MAX_SIZE = 1600

    # AI Icon Positioning
    AI_ICON_POSITION_X_RATIO = 7 / 8
    AI_ICON_SCALE_RATIO = 1 / 8

    # Shape Fitting Constants
    SHAPE_MARGIN_IN = 0.05
    MIN_CONTENT_SIZE_IN = 1.0
    MIN_BBOX_SIZE = 1e-6
    MIN_USABLE_AREA_IN = 0.2
    MARGIN_MULTIPLIER = 2
    CENTER_DIVISOR = 2.0

    # Slide Shape Names (template placeholders)
    SHAPE_NAME_TITLE = "title"
    SHAPE_NAME_LEAD = "lead"
    SHAPE_NAME_FOOTER = "footer"

    # JSON Data Keys
    JSON_KEY_BODY = "body"
    JSON_KEY_TITLE = "title"
    JSON_KEY_LEAD = "lead"

    # Make Types
    MAKE_TYPE_OCR = "ocr"

    # Image Format
    IMAGE_FORMAT_PNG = "PNG"

    # Default Positions
    DEFAULT_SLIDE_INDEX = 0
    AI_ICON_Y_POSITION = 0

    # OCR Tag Values
    OCR_TAG_TITLE = "Title"
    OCR_TAG_LEAD = "Lead"

    # Fallback Textbox Layout
    FALLBACK_PPI = 102.4
    FALLBACK_X = 0.7
    FALLBACK_Y = 2.0
    FALLBACK_WIDTH = 12.0
    FALLBACK_HEIGHT = 5.0
    FALLBACK_FONT_SIZE = 14
    FALLBACK_FONT_COLOR = (0, 0, 0)
    FALLBACK_TEXT_ALIGN = "left"
    FALLBACK_VERTICAL_ALIGN = "top"

    # Shape Border
    DEFAULT_BORDER_WITH_LINE = 2.25
    DEFAULT_BORDER_WITHOUT_LINE = 0

    # File Upload Constraints
    ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}
    MAX_FILE_SIZE_MB = 10
    MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

    # List-style text detection
    LIST_BULLET_CHARS = frozenset('・●○◆◇■□▶▷▸▹►•▪▫◉◎※→➔➡✓✔−－')
    LIST_NUMBERED_PATTERN = re.compile(
        r'^(\d+[.．、）)]\s|'
        r'[①-⑳]\s?|'
        r'[⓪-⓿]\s?|'
        r'（\d+[）)]\s?)'
    )

    # PPI for Gemini 1K image
    PPI_1K = 102.4


class GEMINIConfig:
    MAX_RETRIES = 10
    BASE_DELAY = 1.0
