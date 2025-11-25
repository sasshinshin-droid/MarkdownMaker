#!/usr/bin/env python3
"""
Command-line interface for Enhanced MarkdownMaker
Provides user-friendly CLI for PowerPoint to Markdown conversion
"""

import argparse
import sys
import logging
from pathlib import Path
from typing import Optional
import json

from enhanced_markdown_maker import EnhancedMarkdownMaker
from image_processor import ImageProcessor
from mermaid_generator import MermaidGenerator


# ASCII art banner
BANNER = """
╔═══════════════════════════════════════════════════════════╗
║                                                           ║
║        Enhanced MarkdownMaker                             ║
║        PowerPoint to Markdown Converter                   ║
║                                                           ║
║        Built on Microsoft MarkItDown                      ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝
"""


def setup_logging(verbose: bool = False, log_file: Optional[str] = None):
    """Setup logging configuration"""
    level = logging.DEBUG if verbose else logging.INFO

    handlers = [logging.StreamHandler()]
    if log_file:
        handlers.append(logging.FileHandler(log_file))

    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )


def convert_command(args):
    """Handle the convert command"""
    print(BANNER)
    print(f"📄 Converting: {args.input}")
    print(f"📁 Output directory: {args.output_dir}")
    print(f"🔢 Object threshold: {args.threshold}")
    print("-" * 60)

    try:
        # Initialize the converter
        maker = EnhancedMarkdownMaker(
            llm_client=None,  # Could be configured from args
            output_dir=args.output_dir,
            object_threshold=args.threshold,
            enable_mermaid=args.enable_mermaid
        )

        # Perform conversion
        markdown, metadata = maker.convert_powerpoint(
            args.input,
            output_filename=args.output_name
        )

        # Display results
        print("\n✅ Conversion complete!")
        print(f"📊 Total slides: {metadata.total_slides}")
        print(f"🖼️  Slides saved as images: {metadata.slides_as_images}")
        print(f"📝 Slides saved as text: {metadata.slides_as_text}")
        print(f"⏱️  Conversion time: {metadata.conversion_time:.2f}s")

        if args.enable_mermaid:
            print(f"📊 Mermaid diagrams: {metadata.mermaid_diagrams}")

        output_file = Path(args.output_dir) / (args.output_name or f"{Path(args.input).stem}.md")
        print(f"\n📄 Markdown file: {output_file}")
        print(f"📁 Slide images: {Path(args.output_dir) / 'slides'}")

        if args.enable_mermaid:
            print(f"📊 Mermaid diagrams: {Path(args.output_dir) / 'mermaid'}")

        # Show slide breakdown
        if args.verbose:
            print("\n📋 Slide Breakdown:")
            for analysis in metadata.slide_analyses:
                print(f"  Slide {analysis.slide_number}: "
                      f"{analysis.object_count} objects, "
                      f"complexity={analysis.complexity_score:.1f} "
                      f"{'[IMAGE]' if analysis.should_save_as_image else '[TEXT]'}")

        return 0

    except Exception as e:
        print(f"\n❌ Error during conversion: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def batch_command(args):
    """Handle the batch convert command"""
    print(BANNER)
    print(f"📦 Batch converting files from: {args.input_dir}")
    print(f"📁 Output directory: {args.output_dir}")
    print("-" * 60)

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"❌ Error: {input_dir} is not a directory")
        return 1

    # Find all PowerPoint files
    pptx_files = list(input_dir.glob("*.pptx")) + list(input_dir.glob("*.ppt"))

    if not pptx_files:
        print(f"❌ No PowerPoint files found in {input_dir}")
        return 1

    print(f"Found {len(pptx_files)} PowerPoint file(s)")
    print()

    # Initialize converter
    maker = EnhancedMarkdownMaker(
        llm_client=None,
        output_dir=args.output_dir,
        object_threshold=args.threshold,
        enable_mermaid=args.enable_mermaid
    )

    # Process each file
    results = []
    for i, pptx_file in enumerate(pptx_files, 1):
        print(f"[{i}/{len(pptx_files)}] Converting: {pptx_file.name}")

        try:
            markdown, metadata = maker.convert_powerpoint(str(pptx_file))
            results.append({
                "file": pptx_file.name,
                "success": True,
                "slides": metadata.total_slides,
                "time": metadata.conversion_time
            })
            print(f"  ✅ Success: {metadata.total_slides} slides in {metadata.conversion_time:.2f}s")

        except Exception as e:
            results.append({
                "file": pptx_file.name,
                "success": False,
                "error": str(e)
            })
            print(f"  ❌ Failed: {e}")

        print()

    # Summary
    successful = sum(1 for r in results if r["success"])
    failed = len(results) - successful

    print("=" * 60)
    print(f"✅ Batch conversion complete!")
    print(f"  Success: {successful}/{len(results)}")
    print(f"  Failed: {failed}/{len(results)}")

    # Save batch report
    if args.report:
        report_path = Path(args.output_dir) / "batch_report.json"
        with open(report_path, 'w') as f:
            json.dump(results, f, indent=2)
        print(f"\n📄 Report saved to: {report_path}")

    return 0 if failed == 0 else 1


def analyze_command(args):
    """Handle the analyze command (analyze without full conversion)"""
    print(BANNER)
    print(f"🔍 Analyzing: {args.input}")
    print("-" * 60)

    try:
        from pptx import Presentation

        prs = Presentation(args.input)

        print(f"\n📊 Presentation Analysis:")
        print(f"  Total slides: {len(prs.slides)}")
        print(f"  Slide dimensions: {prs.slide_width} x {prs.slide_height}")
        print()

        # Analyze each slide
        maker = EnhancedMarkdownMaker(
            llm_client=None,
            output_dir="temp",
            object_threshold=args.threshold
        )

        analyses = maker._analyze_all_slides(prs)

        print("📋 Slide-by-slide Analysis:")
        print()

        for analysis in analyses:
            print(f"Slide {analysis.slide_number}: {analysis.title or '(No title)'}")
            print(f"  Objects: {analysis.object_count}")
            print(f"  Complexity: {analysis.complexity_score:.1f}")
            print(f"  Has image: {analysis.has_image}")
            print(f"  Has chart: {analysis.has_chart}")
            print(f"  Has table: {analysis.has_table}")
            print(f"  Should save as image: {analysis.should_save_as_image}")

            if args.verbose:
                print("  Object breakdown:")
                for obj in analysis.objects[:5]:
                    print(f"    - {obj.object_type}: {obj.description[:50]}")
                if len(analysis.objects) > 5:
                    print(f"    ... and {len(analysis.objects) - 5} more")

            print()

        # Statistics
        total_objects = sum(a.object_count for a in analyses)
        avg_objects = total_objects / len(analyses) if analyses else 0
        image_slides = sum(1 for a in analyses if a.should_save_as_image)

        print("=" * 60)
        print("📊 Statistics:")
        print(f"  Total objects: {total_objects}")
        print(f"  Average objects per slide: {avg_objects:.1f}")
        print(f"  Slides with {args.threshold}+ objects: {image_slides}")
        print(f"  Slides with images: {sum(1 for a in analyses if a.has_image)}")
        print(f"  Slides with charts: {sum(1 for a in analyses if a.has_chart)}")
        print(f"  Slides with tables: {sum(1 for a in analyses if a.has_table)}")

        return 0

    except Exception as e:
        print(f"\n❌ Error during analysis: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def mermaid_command(args):
    """Handle the mermaid generation command"""
    print(BANNER)
    print(f"📊 Generating Mermaid diagrams from: {args.input}")
    print("-" * 60)

    try:
        generator = MermaidGenerator()

        # Create sample data (in production, this would come from LLM analysis)
        sample_analysis = {
            "objects": [
                {"type": "shape", "description": "Start"},
                {"type": "shape", "description": "Process Data"},
                {"type": "shape", "description": "End"}
            ],
            "slide_type": args.diagram_type,
            "description": "Sample diagram",
            "title": Path(args.input).stem
        }

        mermaid_code = generator.generate_from_analysis(sample_analysis)

        if mermaid_code:
            print("\n✅ Mermaid diagram generated!")
            print("\n" + "=" * 60)
            print(mermaid_code)
            print("=" * 60)

            if args.output:
                output_path = Path(args.output)
                generator.save_mermaid_file(mermaid_code, output_path)
                print(f"\n📄 Saved to: {output_path}")

            return 0
        else:
            print("❌ Could not generate Mermaid diagram")
            return 1

    except Exception as e:
        print(f"\n❌ Error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def create_parser() -> argparse.ArgumentParser:
    """Create the argument parser"""
    parser = argparse.ArgumentParser(
        description="Enhanced MarkdownMaker - Convert PowerPoint to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert a single PowerPoint file
  %(prog)s convert presentation.pptx

  # Convert with custom output directory
  %(prog)s convert presentation.pptx -o output/docs

  # Batch convert all files in a directory
  %(prog)s batch input_dir/ -o output/

  # Analyze a presentation without converting
  %(prog)s analyze presentation.pptx -v

  # Generate Mermaid diagram
  %(prog)s mermaid presentation.pptx -t flowchart -o diagram.mmd

For more information, visit: https://github.com/microsoft/markitdown
        """
    )

    parser.add_argument(
        '--version',
        action='version',
        version='Enhanced MarkdownMaker v1.0.0'
    )

    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # Convert command
    convert_parser = subparsers.add_parser('convert', help='Convert a PowerPoint file to Markdown')
    convert_parser.add_argument('input', help='Input PowerPoint file (.pptx or .ppt)')
    convert_parser.add_argument('-o', '--output-dir', default='output', help='Output directory')
    convert_parser.add_argument('-n', '--output-name', help='Output filename')
    convert_parser.add_argument('-t', '--threshold', type=int, default=5,
                              help='Object count threshold for image conversion (default: 5)')
    convert_parser.add_argument('-m', '--enable-mermaid', action='store_true',
                              help='Enable Mermaid diagram generation')
    convert_parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')
    convert_parser.add_argument('--log-file', help='Log file path')

    # Batch command
    batch_parser = subparsers.add_parser('batch', help='Batch convert multiple PowerPoint files')
    batch_parser.add_argument('input_dir', help='Input directory containing PowerPoint files')
    batch_parser.add_argument('-o', '--output-dir', default='output', help='Output directory')
    batch_parser.add_argument('-t', '--threshold', type=int, default=5,
                            help='Object count threshold (default: 5)')
    batch_parser.add_argument('-m', '--enable-mermaid', action='store_true',
                            help='Enable Mermaid diagram generation')
    batch_parser.add_argument('-r', '--report', action='store_true',
                            help='Generate batch conversion report')
    batch_parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')

    # Analyze command
    analyze_parser = subparsers.add_parser('analyze', help='Analyze a PowerPoint file without converting')
    analyze_parser.add_argument('input', help='Input PowerPoint file')
    analyze_parser.add_argument('-t', '--threshold', type=int, default=5,
                              help='Object count threshold (default: 5)')
    analyze_parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')

    # Mermaid command
    mermaid_parser = subparsers.add_parser('mermaid', help='Generate Mermaid diagram from slide')
    mermaid_parser.add_argument('input', help='Input PowerPoint file')
    mermaid_parser.add_argument('-t', '--diagram-type', default='flowchart',
                              choices=['flowchart', 'sequence', 'class', 'state', 'pie', 'mindmap'],
                              help='Diagram type to generate')
    mermaid_parser.add_argument('-o', '--output', help='Output .mmd file')
    mermaid_parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')

    return parser


def main():
    """Main entry point"""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 0

    # Setup logging
    setup_logging(
        verbose=getattr(args, 'verbose', False),
        log_file=getattr(args, 'log_file', None)
    )

    # Route to appropriate command handler
    command_handlers = {
        'convert': convert_command,
        'batch': batch_command,
        'analyze': analyze_command,
        'mermaid': mermaid_command
    }

    handler = command_handlers.get(args.command)
    if handler:
        return handler(args)
    else:
        print(f"Unknown command: {args.command}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
