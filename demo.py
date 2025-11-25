#!/usr/bin/env python3
"""
Complete demonstration of Enhanced MarkdownMaker system
Shows all major features and capabilities
"""

import sys
from pathlib import Path
import json

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from enhanced_markdown_maker import EnhancedMarkdownMaker
from image_processor import ImageProcessor
from mermaid_generator import MermaidGenerator, DiagramType
from config import ConfigManager, create_default_config, print_config


BANNER = """
╔═══════════════════════════════════════════════════════════╗
║                                                           ║
║        Enhanced MarkdownMaker - Complete Demo            ║
║        PowerPoint to Markdown Conversion System          ║
║                                                           ║
║        Built on Microsoft MarkItDown                      ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝
"""


def demo_1_system_overview():
    """Demo 1: System overview and capabilities"""
    print("\n" + "=" * 70)
    print("DEMO 1: System Overview")
    print("=" * 70)

    print("\n📋 System Capabilities:")
    print("  ✓ PowerPoint to Markdown conversion")
    print("  ✓ Automatic object counting per slide")
    print("  ✓ Smart image generation for complex slides")
    print("  ✓ Mermaid diagram generation")
    print("  ✓ LLM integration for advanced analysis")
    print("  ✓ Batch processing support")
    print("  ✓ Flexible configuration management")

    print("\n🏗️  Architecture:")
    print("  • EnhancedMarkdownMaker - Main converter")
    print("  • ImageProcessor - Image extraction & processing")
    print("  • MermaidGenerator - Diagram generation")
    print("  • ConfigManager - Configuration handling")
    print("  • CLI - Command-line interface")

    print("\n📦 Dependencies:")
    print("  • markitdown >= 0.1.3")
    print("  • python-pptx >= 0.6.21")
    print("  • Pillow >= 9.0.0")
    print("  • pyyaml >= 6.0.0")
    print("  • openai >= 1.0.0 (optional)")


def demo_2_configuration():
    """Demo 2: Configuration management"""
    print("\n" + "=" * 70)
    print("DEMO 2: Configuration Management")
    print("=" * 70)

    # Create config manager
    config_manager = ConfigManager()

    print("\n📄 Default Configuration:")
    print_config(config_manager.config)

    # Save example config
    example_path = Path("output/demo/config_example.yaml")
    example_path.parent.mkdir(parents=True, exist_ok=True)

    config_manager.save_config(example_path, format='yaml')
    print(f"\n💾 Example config saved to: {example_path}")

    # Demonstrate config updates
    print("\n🔧 Updating configuration...")
    config_manager.update_config({
        'conversion': {
            'object_threshold': 3,
            'enable_mermaid': True
        },
        'image': {
            'width': 1280,
            'height': 720
        }
    })

    print("✓ Configuration updated")
    print(f"  New threshold: {config_manager.config.conversion.object_threshold}")
    print(f"  New image size: {config_manager.config.image.width}x{config_manager.config.image.height}")


def demo_3_basic_conversion():
    """Demo 3: Basic PowerPoint conversion"""
    print("\n" + "=" * 70)
    print("DEMO 3: Basic Conversion (Simulated)")
    print("=" * 70)

    # Initialize converter
    maker = EnhancedMarkdownMaker(
        llm_client=None,
        output_dir="output/demo/basic",
        object_threshold=5,
        enable_mermaid=True
    )

    print("\n✓ EnhancedMarkdownMaker initialized")
    print(f"  Output directory: {maker.output_dir}")
    print(f"  Object threshold: {maker.object_threshold}")
    print(f"  Mermaid enabled: {maker.enable_mermaid}")

    print("\n📊 Conversion workflow:")
    print("  1. Load PowerPoint file")
    print("  2. Analyze each slide for objects")
    print("  3. Count shapes, images, charts, tables")
    print("  4. Determine if slide should be saved as image (≥5 objects)")
    print("  5. Generate Mermaid diagrams for complex visuals")
    print("  6. Create comprehensive Markdown output")
    print("  7. Save metadata and statistics")

    print("\n📝 Example output structure:")
    print("  output/demo/basic/")
    print("    ├── presentation.md")
    print("    ├── slides/")
    print("    │   ├── presentation_slide_001.png")
    print("    │   └── presentation_slide_005.png")
    print("    ├── mermaid/")
    print("    │   ├── slide_001.mmd")
    print("    │   └── slide_005.mmd")
    print("    └── metadata.json")


def demo_4_mermaid_generation():
    """Demo 4: Mermaid diagram generation"""
    print("\n" + "=" * 70)
    print("DEMO 4: Mermaid Diagram Generation")
    print("=" * 70)

    generator = MermaidGenerator(enable_styling=True)

    # Example 1: Flowchart
    print("\n📊 Example 1: Flowchart")
    print("-" * 70)

    flowchart_data = {
        "objects": [
            {"type": "shape", "description": "Start: User Login"},
            {"type": "shape", "description": "Verify Credentials"},
            {"type": "shape", "description": "Check Authorization"},
            {"type": "shape", "description": "Grant Access"},
            {"type": "shape", "description": "End: User Dashboard"}
        ],
        "slide_type": "flowchart",
        "description": "User authentication workflow",
        "title": "Login Process"
    }

    flowchart = generator.generate_from_analysis(flowchart_data)
    print(flowchart)

    # Save
    output_path = Path("output/demo/mermaid/flowchart_example.mmd")
    generator.save_mermaid_file(flowchart, output_path)
    print(f"\n💾 Saved to: {output_path}")

    # Example 2: Sequence Diagram
    print("\n📊 Example 2: Sequence Diagram")
    print("-" * 70)

    sequence_data = {
        "objects": [
            {"type": "text", "description": "Frontend"},
            {"type": "text", "description": "Backend API"},
            {"type": "text", "description": "Database"}
        ],
        "slide_type": "sequence",
        "description": "API request flow",
        "title": "Data Fetch Sequence"
    }

    sequence = generator.generate_from_analysis(sequence_data, DiagramType.SEQUENCE)
    print(sequence)

    # Example 3: Pie Chart
    print("\n📊 Example 3: Pie Chart")
    print("-" * 70)

    pie_data = {
        "objects": [
            {"type": "chart", "description": "Frontend: 40%"},
            {"type": "chart", "description": "Backend: 35%"},
            {"type": "chart", "description": "Database: 25%"}
        ],
        "slide_type": "pie",
        "description": "Resource allocation",
        "title": "Project Time Distribution"
    }

    pie = generator.generate_from_analysis(pie_data, DiagramType.PIE)
    print(pie)


def demo_5_image_processing():
    """Demo 5: Image processing capabilities"""
    print("\n" + "=" * 70)
    print("DEMO 5: Image Processing")
    print("=" * 70)

    processor = ImageProcessor(
        output_dir=Path("output/demo/images"),
        width=1920,
        height=1080,
        dpi=300,
        quality=95
    )

    print("\n✓ ImageProcessor initialized")
    print(f"  Output directory: {processor.output_dir}")
    print(f"  Resolution: {processor.width}x{processor.height}")
    print(f"  DPI: {processor.dpi}")
    print(f"  Quality: {processor.quality}")

    print("\n🖼️  Image Processing Features:")
    print("  • Slide extraction using LibreOffice/unoconv")
    print("  • Fallback to placeholder generation")
    print("  • Automatic image optimization")
    print("  • Thumbnail creation")
    print("  • Watermark support")
    print("  • Batch processing")

    print("\n📏 Image Specifications:")
    print(f"  • Default size: {processor.width}x{processor.height}")
    print(f"  • DPI: {processor.dpi}")
    print(f"  • Thumbnail size: {processor.THUMBNAIL_SIZE}")
    print(f"  • Format: PNG")

    print("\n🔧 Available Methods:")
    print("  • extract_slide_image() - Extract single slide")
    print("  • create_thumbnail() - Generate thumbnail")
    print("  • add_watermark() - Add watermark to image")
    print("  • batch_process() - Process multiple slides")


def demo_6_cli_usage():
    """Demo 6: CLI usage examples"""
    print("\n" + "=" * 70)
    print("DEMO 6: Command-Line Interface")
    print("=" * 70)

    print("\n💻 Basic Commands:")
    print("\n1. Convert a single file:")
    print("   $ markdown-maker convert presentation.pptx")

    print("\n2. Convert with options:")
    print("   $ markdown-maker convert presentation.pptx \\")
    print("       --output-dir ./docs \\")
    print("       --threshold 3 \\")
    print("       --enable-mermaid \\")
    print("       --verbose")

    print("\n3. Batch conversion:")
    print("   $ markdown-maker batch ./presentations/ \\")
    print("       --output-dir ./output \\")
    print("       --report")

    print("\n4. Analyze without converting:")
    print("   $ markdown-maker analyze presentation.pptx --verbose")

    print("\n5. Generate Mermaid diagram:")
    print("   $ markdown-maker mermaid presentation.pptx \\")
    print("       --diagram-type flowchart \\")
    print("       --output diagram.mmd")

    print("\n📋 Available Commands:")
    print("  • convert  - Convert PowerPoint to Markdown")
    print("  • batch    - Batch convert multiple files")
    print("  • analyze  - Analyze presentation structure")
    print("  • mermaid  - Generate Mermaid diagrams")

    print("\n⚙️  Common Options:")
    print("  -o, --output-dir    Output directory")
    print("  -t, --threshold     Object count threshold")
    print("  -m, --enable-mermaid Enable Mermaid generation")
    print("  -v, --verbose       Verbose output")
    print("  --log-file          Log file path")


def demo_7_output_examples():
    """Demo 7: Output format examples"""
    print("\n" + "=" * 70)
    print("DEMO 7: Output Format Examples")
    print("=" * 70)

    print("\n📄 Markdown Output Structure:")
    print("""
# Presentation Title

*Converted from PowerPoint on 2025-11-25 12:00:00*
**Total Slides:** 10

## Table of Contents

- [Slide 1: Introduction](#slide-1)
- [Slide 2: Overview](#slide-2)
- [Slide 3: Architecture](#slide-3)
...

---

## Slide 1

### Introduction

**Objects:** 3 | **Complexity:** 5.0

Welcome to the Enhanced MarkdownMaker demonstration.
This tool converts PowerPoint presentations to Markdown format.

**Slide Elements:**
- text: Title Text
- shape: Background Box
- image: Company Logo

---

## Slide 2

### System Architecture

**Objects:** 8 | **Complexity:** 12.0

![Slide 2](slides/presentation_slide_002.png)

### Diagram Representation

```mermaid
flowchart TD
    A[Input] --> B[Processing]
    B --> C[Output]
```

**Slide Elements:**
- shape: Component 1
- shape: Component 2
- shape: Component 3
...
    """)

    print("\n📊 Metadata JSON Structure:")
    metadata_example = {
        "source_file": "presentation.pptx",
        "total_slides": 10,
        "slides_as_images": 4,
        "slides_as_text": 6,
        "mermaid_diagrams": 3,
        "conversion_time": 8.5,
        "timestamp": "2025-11-25T12:00:00",
        "slide_analyses": "..."
    }

    print(json.dumps(metadata_example, indent=2))


def demo_8_advanced_features():
    """Demo 8: Advanced features"""
    print("\n" + "=" * 70)
    print("DEMO 8: Advanced Features")
    print("=" * 70)

    print("\n🤖 LLM Integration:")
    print("  • Automatic object detection from images")
    print("  • Intelligent diagram type classification")
    print("  • Semantic content analysis")
    print("  • Context-aware descriptions")
    print("  • Smart Mermaid generation")

    print("\n🔧 Extensibility:")
    print("  • Plugin architecture for custom processors")
    print("  • Custom output format support")
    print("  • Webhook integration for automation")
    print("  • API endpoints for web integration")

    print("\n⚡ Performance Optimizations:")
    print("  • Parallel slide processing")
    print("  • Image caching")
    print("  • Incremental conversion")
    print("  • Memory-efficient streaming")

    print("\n🔐 Enterprise Features:")
    print("  • Watermarking support")
    print("  • Access control integration")
    print("  • Audit logging")
    print("  • Batch reporting")


def main():
    """Run complete demonstration"""
    print(BANNER)

    demos = [
        ("System Overview", demo_1_system_overview),
        ("Configuration Management", demo_2_configuration),
        ("Basic Conversion", demo_3_basic_conversion),
        ("Mermaid Generation", demo_4_mermaid_generation),
        ("Image Processing", demo_5_image_processing),
        ("CLI Usage", demo_6_cli_usage),
        ("Output Examples", demo_7_output_examples),
        ("Advanced Features", demo_8_advanced_features),
    ]

    for name, demo_func in demos:
        try:
            demo_func()
        except Exception as e:
            print(f"\n❌ Error in {name}: {e}")
            import traceback
            traceback.print_exc()

    # Summary
    print("\n" + "=" * 70)
    print("🎉 Demonstration Complete!")
    print("=" * 70)

    print("\n📚 Next Steps:")
    print("  1. Install dependencies: pip install -r requirements.txt")
    print("  2. Try the examples: python examples/basic_usage.py")
    print("  3. Run the CLI: markdown-maker --help")
    print("  4. Read the documentation: README.md")
    print("  5. Run tests: pytest tests/")

    print("\n🔗 Resources:")
    print("  • GitHub: https://github.com/microsoft/markitdown")
    print("  • Documentation: See README.md")
    print("  • Examples: See examples/ directory")
    print("  • Tests: See tests/ directory")

    print("\n✨ Thank you for using Enhanced MarkdownMaker!")
    print()


if __name__ == "__main__":
    main()
