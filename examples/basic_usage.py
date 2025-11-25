#!/usr/bin/env python3
"""
Basic usage examples for Enhanced MarkdownMaker
Demonstrates common use cases and features
"""

import sys
from pathlib import Path

# Add src to path for direct execution
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from enhanced_markdown_maker import EnhancedMarkdownMaker
from image_processor import ImageProcessor
from mermaid_generator import MermaidGenerator, DiagramType
from config import ConfigManager


def example_1_basic_conversion():
    """Example 1: Basic PowerPoint to Markdown conversion"""
    print("=" * 60)
    print("Example 1: Basic Conversion")
    print("=" * 60)

    # Initialize the converter
    maker = EnhancedMarkdownMaker(
        llm_client=None,  # No LLM client for basic usage
        output_dir="output/example1",
        object_threshold=5
    )

    # Convert a PowerPoint file (you'll need to provide your own .pptx file)
    # markdown, metadata = maker.convert_powerpoint("sample.pptx")

    print("\nInitialized EnhancedMarkdownMaker")
    print(f"  Output directory: output/example1")
    print(f"  Object threshold: 5")
    print(f"  Mermaid enabled: True")

    # In production, you would call:
    # markdown, metadata = maker.convert_powerpoint("your_presentation.pptx")
    # print(f"\nConversion complete!")
    # print(f"  Total slides: {metadata.total_slides}")
    # print(f"  Slides as images: {metadata.slides_as_images}")


def example_2_custom_configuration():
    """Example 2: Using custom configuration"""
    print("\n" + "=" * 60)
    print("Example 2: Custom Configuration")
    print("=" * 60)

    # Load configuration
    config_manager = ConfigManager()

    # Update configuration
    config_manager.update_config({
        'conversion': {
            'object_threshold': 3,  # Lower threshold
            'enable_mermaid': True,
            'watermark_images': True
        },
        'image': {
            'width': 1280,
            'height': 720,
            'dpi': 200
        }
    })

    # Use configuration
    config = config_manager.get_conversion_config()
    print(f"\nCustom configuration:")
    print(f"  Object threshold: {config.object_threshold}")
    print(f"  Enable Mermaid: {config.enable_mermaid}")
    print(f"  Watermark images: {config.watermark_images}")

    # Create converter with custom config
    maker = EnhancedMarkdownMaker(
        llm_client=None,
        output_dir=config.output_dir,
        object_threshold=config.object_threshold,
        enable_mermaid=config.enable_mermaid
    )

    print("\nConverter initialized with custom config")


def example_3_image_processing():
    """Example 3: Standalone image processing"""
    print("\n" + "=" * 60)
    print("Example 3: Image Processing")
    print("=" * 60)

    # Initialize image processor
    processor = ImageProcessor(
        output_dir=Path("output/example3/images"),
        width=1920,
        height=1080,
        dpi=300
    )

    print("\nImageProcessor initialized:")
    print(f"  Output directory: output/example3/images")
    print(f"  Resolution: 1920x1080")
    print(f"  DPI: 300")

    # In production, you would use:
    # image_path = processor.extract_slide_image(
    #     pptx_path=Path("presentation.pptx"),
    #     slide_number=1
    # )
    #
    # # Create thumbnail
    # thumbnail_path = processor.create_thumbnail(image_path)
    #
    # # Add watermark
    # watermarked = processor.add_watermark(
    #     image_path,
    #     watermark_text="Confidential",
    #     position="bottom-right"
    # )

    print("\nAvailable methods:")
    print("  - extract_slide_image()")
    print("  - create_thumbnail()")
    print("  - add_watermark()")
    print("  - batch_process()")


def example_4_mermaid_generation():
    """Example 4: Generate Mermaid diagrams"""
    print("\n" + "=" * 60)
    print("Example 4: Mermaid Diagram Generation")
    print("=" * 60)

    # Initialize Mermaid generator
    generator = MermaidGenerator(enable_styling=True)

    # Example 1: Flowchart
    flowchart_data = {
        "objects": [
            {"type": "shape", "description": "Start Process"},
            {"type": "shape", "description": "Collect Input"},
            {"type": "shape", "description": "Validate Data"},
            {"type": "shape", "description": "Process Results"},
            {"type": "shape", "description": "End Process"}
        ],
        "slide_type": "flowchart",
        "description": "Data processing workflow",
        "title": "Process Flow"
    }

    mermaid_code = generator.generate_from_analysis(flowchart_data)

    print("\nGenerated Flowchart:")
    print("-" * 60)
    print(mermaid_code)
    print("-" * 60)

    # Save to file
    output_path = Path("output/example4/flowchart.mmd")
    generator.save_mermaid_file(mermaid_code, output_path)
    print(f"\nSaved to: {output_path}")

    # Example 2: Sequence Diagram
    sequence_data = {
        "objects": [
            {"type": "text", "description": "Client"},
            {"type": "text", "description": "API Server"},
            {"type": "text", "description": "Database"}
        ],
        "slide_type": "sequence",
        "description": "API interaction flow",
        "title": "API Sequence"
    }

    sequence_code = generator.generate_from_analysis(
        sequence_data,
        diagram_type=DiagramType.SEQUENCE
    )

    print("\nGenerated Sequence Diagram:")
    print("-" * 60)
    print(sequence_code)
    print("-" * 60)


def example_5_batch_conversion():
    """Example 5: Batch processing multiple files"""
    print("\n" + "=" * 60)
    print("Example 5: Batch Conversion")
    print("=" * 60)

    # Initialize converter
    maker = EnhancedMarkdownMaker(
        llm_client=None,
        output_dir="output/example5",
        object_threshold=5
    )

    # Simulate batch processing
    presentations = [
        "presentation1.pptx",
        "presentation2.pptx",
        "presentation3.pptx"
    ]

    print("\nBatch processing setup:")
    print(f"  Input files: {len(presentations)}")
    print(f"  Output directory: output/example5")

    print("\nIn production, you would:")
    print("  1. Loop through all .pptx files")
    print("  2. Convert each file")
    print("  3. Collect statistics")
    print("  4. Generate batch report")

    # In production:
    # results = []
    # for pptx_file in presentations:
    #     try:
    #         markdown, metadata = maker.convert_powerpoint(pptx_file)
    #         results.append({
    #             "file": pptx_file,
    #             "success": True,
    #             "slides": metadata.total_slides
    #         })
    #     except Exception as e:
    #         results.append({
    #             "file": pptx_file,
    #             "success": False,
    #             "error": str(e)
    #         })


def example_6_with_llm():
    """Example 6: Using LLM integration for advanced analysis"""
    print("\n" + "=" * 60)
    print("Example 6: LLM Integration")
    print("=" * 60)

    print("\nLLM integration allows:")
    print("  - Automatic object counting from slide images")
    print("  - Intelligent diagram type detection")
    print("  - Semantic analysis of slide content")
    print("  - Auto-generated Mermaid diagrams")

    print("\nTo enable LLM integration:")
    print("  1. Install OpenAI: pip install openai")
    print("  2. Set API key: export OPENAI_API_KEY=your-key")
    print("  3. Initialize with LLM client:")

    print("""
    from openai import OpenAI

    client = OpenAI(api_key="your-api-key")

    maker = EnhancedMarkdownMaker(
        llm_client=client,
        llm_model="gpt-4o",
        output_dir="output/with_llm"
    )

    # Now conversion will use LLM for enhanced analysis
    markdown, metadata = maker.convert_powerpoint("presentation.pptx")
    """)


def example_7_configuration_files():
    """Example 7: Using configuration files"""
    print("\n" + "=" * 60)
    print("Example 7: Configuration Files")
    print("=" * 60)

    # Create example configuration
    from config import create_default_config

    config_path = Path("output/example7/markdown_maker.yaml")
    config_path.parent.mkdir(parents=True, exist_ok=True)

    create_default_config(config_path, format='yaml')

    print(f"\nCreated example config: {config_path}")
    print("\nConfiguration file allows:")
    print("  - Persistent settings across sessions")
    print("  - Easy sharing of configurations")
    print("  - Environment-specific settings")
    print("  - No need to pass parameters every time")

    print("\nUsage with config file:")
    print("  manager = ConfigManager('markdown_maker.yaml')")
    print("  config = manager.get_conversion_config()")
    print("  maker = EnhancedMarkdownMaker(**config)")


def main():
    """Run all examples"""
    print("\n")
    print("╔═══════════════════════════════════════════════════════════╗")
    print("║                                                           ║")
    print("║        Enhanced MarkdownMaker - Usage Examples           ║")
    print("║                                                           ║")
    print("╚═══════════════════════════════════════════════════════════╝")
    print("\n")

    examples = [
        example_1_basic_conversion,
        example_2_custom_configuration,
        example_3_image_processing,
        example_4_mermaid_generation,
        example_5_batch_conversion,
        example_6_with_llm,
        example_7_configuration_files
    ]

    for example in examples:
        try:
            example()
        except Exception as e:
            print(f"\n❌ Error in {example.__name__}: {e}")

    print("\n" + "=" * 60)
    print("Examples complete!")
    print("=" * 60)
    print("\nFor more information:")
    print("  - See README.md for full documentation")
    print("  - Run: markdown-maker --help")
    print("  - Visit: https://github.com/microsoft/markitdown")
    print()


if __name__ == "__main__":
    main()
