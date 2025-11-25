#!/usr/bin/env python3
"""
Image Processing Module for Enhanced MarkdownMaker
Handles slide image extraction, optimization, and manipulation
"""

import logging
from pathlib import Path
from typing import Optional, Tuple, Dict
import tempfile
import subprocess

try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    import io
except ImportError as e:
    raise ImportError(f"Missing required dependency: {e}. Please install: pip install Pillow")


logger = logging.getLogger(__name__)


class ImageProcessor:
    """Handles image processing for PowerPoint slides"""

    DEFAULT_WIDTH = 1920
    DEFAULT_HEIGHT = 1080
    DEFAULT_DPI = 300
    THUMBNAIL_SIZE = (400, 300)

    def __init__(
        self,
        output_dir: Path,
        width: int = DEFAULT_WIDTH,
        height: int = DEFAULT_HEIGHT,
        dpi: int = DEFAULT_DPI,
        quality: int = 95
    ):
        """
        Initialize the ImageProcessor.

        Args:
            output_dir: Directory to save processed images
            width: Output image width in pixels
            height: Output image height in pixels
            dpi: Image DPI for quality
            quality: JPEG quality (1-100)
        """
        self.output_dir = Path(output_dir)
        self.width = width
        self.height = height
        self.dpi = dpi
        self.quality = quality

        self.output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"ImageProcessor initialized: {width}x{height} @ {dpi} DPI")

    def extract_slide_image(
        self,
        pptx_path: Path,
        slide_number: int,
        output_filename: Optional[str] = None
    ) -> Path:
        """
        Extract a single slide as an image using external tools.

        Args:
            pptx_path: Path to PowerPoint file
            slide_number: Slide number (1-indexed)
            output_filename: Optional custom filename

        Returns:
            Path to extracted image
        """
        if output_filename is None:
            output_filename = f"{pptx_path.stem}_slide_{slide_number:03d}.png"

        output_path = self.output_dir / output_filename

        # Try different extraction methods in order of preference
        methods = [
            self._extract_with_libreoffice,
            self._extract_with_unoconv,
            self._create_placeholder_image
        ]

        for method in methods:
            try:
                logger.info(f"Attempting extraction with {method.__name__}")
                result = method(pptx_path, slide_number, output_path)
                if result and result.exists():
                    logger.info(f"Successfully extracted slide {slide_number} to {output_path}")
                    return result
            except Exception as e:
                logger.warning(f"{method.__name__} failed: {e}")
                continue

        # Fallback to placeholder
        logger.warning(f"All extraction methods failed, using placeholder for slide {slide_number}")
        return self._create_placeholder_image(pptx_path, slide_number, output_path)

    def _extract_with_libreoffice(
        self,
        pptx_path: Path,
        slide_number: int,
        output_path: Path
    ) -> Optional[Path]:
        """Extract slide using LibreOffice headless mode"""
        try:
            # Create temporary directory for extraction
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)

                # Convert entire presentation to images
                cmd = [
                    'libreoffice',
                    '--headless',
                    '--convert-to', 'png',
                    '--outdir', str(tmpdir_path),
                    str(pptx_path)
                ]

                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=30
                )

                if result.returncode != 0:
                    logger.error(f"LibreOffice conversion failed: {result.stderr}")
                    return None

                # Find the extracted image for the specific slide
                images = sorted(tmpdir_path.glob("*.png"))
                if slide_number <= len(images):
                    target_image = images[slide_number - 1]
                    # Copy to output location
                    img = Image.open(target_image)
                    img = self._optimize_image(img)
                    img.save(output_path, 'PNG', dpi=(self.dpi, self.dpi))
                    return output_path

                logger.warning(f"Slide {slide_number} not found in extracted images")
                return None

        except FileNotFoundError:
            logger.warning("LibreOffice not found in system PATH")
            return None
        except subprocess.TimeoutExpired:
            logger.error("LibreOffice conversion timed out")
            return None
        except Exception as e:
            logger.error(f"LibreOffice extraction error: {e}")
            return None

    def _extract_with_unoconv(
        self,
        pptx_path: Path,
        slide_number: int,
        output_path: Path
    ) -> Optional[Path]:
        """Extract slide using unoconv tool"""
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)

                cmd = [
                    'unoconv',
                    '-f', 'png',
                    '-o', str(tmpdir_path),
                    str(pptx_path)
                ]

                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=30
                )

                if result.returncode != 0:
                    return None

                images = sorted(tmpdir_path.glob("*.png"))
                if slide_number <= len(images):
                    target_image = images[slide_number - 1]
                    img = Image.open(target_image)
                    img = self._optimize_image(img)
                    img.save(output_path, 'PNG', dpi=(self.dpi, self.dpi))
                    return output_path

                return None

        except FileNotFoundError:
            logger.warning("unoconv not found in system PATH")
            return None
        except Exception as e:
            logger.error(f"unoconv extraction error: {e}")
            return None

    def _create_placeholder_image(
        self,
        pptx_path: Path,
        slide_number: int,
        output_path: Path
    ) -> Path:
        """Create a placeholder image when extraction fails"""
        img = Image.new('RGB', (self.width, self.height), color='#f5f5f5')
        draw = ImageDraw.Draw(img)

        # Try to use a nice font
        try:
            font_large = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 72)
            font_medium = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 36)
            font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 24)
        except:
            font_large = ImageFont.load_default()
            font_medium = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # Draw border
        border_width = 10
        draw.rectangle(
            [border_width, border_width, self.width - border_width, self.height - border_width],
            outline='#cccccc',
            width=border_width
        )

        # Draw text
        y_offset = self.height // 3

        slide_text = f"Slide {slide_number}"
        bbox = draw.textbbox((0, 0), slide_text, font=font_large)
        text_width = bbox[2] - bbox[0]
        draw.text(
            ((self.width - text_width) // 2, y_offset),
            slide_text,
            fill='#333333',
            font=font_large
        )

        subtitle = f"From: {pptx_path.name}"
        bbox = draw.textbbox((0, 0), subtitle, font=font_medium)
        text_width = bbox[2] - bbox[0]
        draw.text(
            ((self.width - text_width) // 2, y_offset + 100),
            subtitle,
            fill='#666666',
            font=font_medium
        )

        note = "Note: Actual slide rendering requires LibreOffice or PowerPoint"
        bbox = draw.textbbox((0, 0), note, font=font_small)
        text_width = bbox[2] - bbox[0]
        draw.text(
            ((self.width - text_width) // 2, y_offset + 180),
            note,
            fill='#999999',
            font=font_small
        )

        img.save(output_path, 'PNG', dpi=(self.dpi, self.dpi))
        return output_path

    def _optimize_image(self, img: Image.Image) -> Image.Image:
        """Optimize image for web display"""
        # Resize if too large
        if img.width > self.width or img.height > self.height:
            img.thumbnail((self.width, self.height), Image.Resampling.LANCZOS)

        # Convert to RGB if necessary
        if img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')

        return img

    def create_thumbnail(
        self,
        image_path: Path,
        thumbnail_size: Optional[Tuple[int, int]] = None
    ) -> Path:
        """
        Create a thumbnail from an image.

        Args:
            image_path: Path to source image
            thumbnail_size: Optional custom size (width, height)

        Returns:
            Path to thumbnail image
        """
        if thumbnail_size is None:
            thumbnail_size = self.THUMBNAIL_SIZE

        img = Image.open(image_path)

        # Create thumbnail
        img.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)

        # Save with _thumb suffix
        thumbnail_path = image_path.parent / f"{image_path.stem}_thumb{image_path.suffix}"
        img.save(thumbnail_path, 'PNG', optimize=True)

        logger.debug(f"Created thumbnail: {thumbnail_path}")
        return thumbnail_path

    def add_watermark(
        self,
        image_path: Path,
        watermark_text: str,
        position: str = "bottom-right",
        opacity: int = 128
    ) -> Path:
        """
        Add watermark to an image.

        Args:
            image_path: Path to source image
            watermark_text: Text to use as watermark
            position: Position of watermark (top-left, top-right, bottom-left, bottom-right, center)
            opacity: Watermark opacity (0-255)

        Returns:
            Path to watermarked image
        """
        img = Image.open(image_path).convert('RGBA')

        # Create watermark layer
        txt_layer = Image.new('RGBA', img.size, (255, 255, 255, 0))
        draw = ImageDraw.Draw(txt_layer)

        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 36)
        except:
            font = ImageFont.load_default()

        # Calculate position
        bbox = draw.textbbox((0, 0), watermark_text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        padding = 20

        if position == "top-left":
            x, y = padding, padding
        elif position == "top-right":
            x, y = img.width - text_width - padding, padding
        elif position == "bottom-left":
            x, y = padding, img.height - text_height - padding
        elif position == "bottom-right":
            x, y = img.width - text_width - padding, img.height - text_height - padding
        else:  # center
            x, y = (img.width - text_width) // 2, (img.height - text_height) // 2

        # Draw watermark
        draw.text((x, y), watermark_text, fill=(255, 255, 255, opacity), font=font)

        # Composite
        watermarked = Image.alpha_composite(img, txt_layer)
        watermarked = watermarked.convert('RGB')

        # Save
        watermarked_path = image_path.parent / f"{image_path.stem}_watermarked{image_path.suffix}"
        watermarked.save(watermarked_path, 'PNG', dpi=(self.dpi, self.dpi))

        logger.debug(f"Added watermark to {watermarked_path}")
        return watermarked_path

    def batch_process(
        self,
        pptx_path: Path,
        slide_numbers: list[int]
    ) -> Dict[int, Path]:
        """
        Process multiple slides in batch.

        Args:
            pptx_path: Path to PowerPoint file
            slide_numbers: List of slide numbers to process

        Returns:
            Dictionary mapping slide numbers to image paths
        """
        results = {}

        for slide_num in slide_numbers:
            try:
                image_path = self.extract_slide_image(pptx_path, slide_num)
                results[slide_num] = image_path
            except Exception as e:
                logger.error(f"Failed to process slide {slide_num}: {e}")
                results[slide_num] = None

        logger.info(f"Batch processed {len(results)} slides")
        return results


if __name__ == "__main__":
    # Test the image processor
    processor = ImageProcessor(output_dir=Path("output/slides"))
    print(f"ImageProcessor initialized: {processor.width}x{processor.height}")
    print("Ready for slide image extraction")
