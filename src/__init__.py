"""
Enhanced MarkdownMaker - PowerPoint to Markdown Conversion System
Built on Microsoft MarkItDown library
"""

from .enhanced_markdown_maker import (
    EnhancedMarkdownMaker,
    SlideObject,
    SlideAnalysis,
    ConversionMetadata,
    ObjectAnalyzer
)

from .image_processor import ImageProcessor

from .mermaid_generator import (
    MermaidGenerator,
    DiagramType,
    DiagramElement
)

from .config import (
    ConfigManager,
    MarkdownMakerConfig,
    ConversionConfig,
    ImageConfig,
    MermaidConfig,
    LLMConfig,
    create_default_config,
    print_config
)


__version__ = "1.0.0"
__author__ = "Enhanced MarkdownMaker Contributors"
__license__ = "MIT"

__all__ = [
    # Main converter
    "EnhancedMarkdownMaker",
    "SlideObject",
    "SlideAnalysis",
    "ConversionMetadata",
    "ObjectAnalyzer",

    # Image processing
    "ImageProcessor",

    # Mermaid generation
    "MermaidGenerator",
    "DiagramType",
    "DiagramElement",

    # Configuration
    "ConfigManager",
    "MarkdownMakerConfig",
    "ConversionConfig",
    "ImageConfig",
    "MermaidConfig",
    "LLMConfig",
    "create_default_config",
    "print_config",
]
