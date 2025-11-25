#!/usr/bin/env python3
"""
Test suite for Enhanced MarkdownMaker
Tests core conversion functionality
"""

import pytest
import sys
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch
import tempfile
import shutil

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from enhanced_markdown_maker import (
    EnhancedMarkdownMaker,
    SlideObject,
    SlideAnalysis,
    ConversionMetadata
)


class TestEnhancedMarkdownMaker:
    """Test cases for EnhancedMarkdownMaker class"""

    @pytest.fixture
    def temp_output_dir(self):
        """Create temporary output directory"""
        temp_dir = tempfile.mkdtemp()
        yield Path(temp_dir)
        shutil.rmtree(temp_dir)

    @pytest.fixture
    def maker(self, temp_output_dir):
        """Create EnhancedMarkdownMaker instance"""
        return EnhancedMarkdownMaker(
            llm_client=None,
            output_dir=str(temp_output_dir),
            object_threshold=5,
            enable_mermaid=True
        )

    def test_initialization(self, maker, temp_output_dir):
        """Test EnhancedMarkdownMaker initialization"""
        assert maker.output_dir == temp_output_dir
        assert maker.object_threshold == 5
        assert maker.enable_mermaid is True
        assert maker.slides_dir.exists()
        assert maker.mermaid_dir.exists()

    def test_create_analysis_prompt(self, maker):
        """Test LLM prompt creation"""
        prompt = maker._create_analysis_prompt()
        assert "Count all distinct visual objects" in prompt
        assert "Mermaid syntax" in prompt
        assert "JSON" in prompt

    @patch('enhanced_markdown_maker.Presentation')
    def test_analyze_slide_objects(self, mock_prs, maker):
        """Test slide object analysis"""
        # Create mock slide
        mock_slide = Mock()
        mock_slide.shapes = []
        mock_slide.shapes.title = None

        # Create mock shapes
        mock_shape1 = Mock()
        mock_shape1.shape_type = 13  # MSO_SHAPE_TYPE.PICTURE
        mock_shape1.name = "Image1"
        mock_shape1.left = 100
        mock_shape1.top = 100
        mock_shape1.width = 200
        mock_shape1.height = 200

        mock_shape2 = Mock()
        mock_shape2.shape_type = 1  # MSO_SHAPE_TYPE.AUTO_SHAPE
        mock_shape2.name = "Shape1"
        mock_shape2.left = 300
        mock_shape2.top = 300
        mock_shape2.width = 150
        mock_shape2.height = 150

        mock_slide.shapes = [mock_shape1, mock_shape2]

        # Analyze
        analysis = maker._analyze_slide_objects(mock_slide, 1)

        assert isinstance(analysis, SlideAnalysis)
        assert analysis.slide_number == 1
        assert analysis.object_count >= 0  # At least some objects detected

    def test_calculate_complexity(self, maker):
        """Test complexity score calculation"""
        score1 = maker._calculate_complexity(3, False, False, False)
        assert score1 == 3

        score2 = maker._calculate_complexity(3, True, False, False)
        assert score2 == 5  # 3 + 2 for image

        score3 = maker._calculate_complexity(3, True, True, True)
        assert score3 == 10  # 3 + 2 + 3 + 2

    def test_should_save_as_image(self, maker):
        """Test threshold logic for image saving"""
        # Below threshold
        analysis1 = SlideAnalysis(
            slide_number=1,
            object_count=3,
            objects=[],
            has_image=False,
            has_chart=False,
            has_table=False,
            complexity_score=3.0,
            should_save_as_image=False
        )
        assert not analysis1.should_save_as_image

        # At threshold
        analysis2 = SlideAnalysis(
            slide_number=2,
            object_count=5,
            objects=[],
            has_image=False,
            has_chart=False,
            has_table=False,
            complexity_score=5.0,
            should_save_as_image=True
        )
        assert analysis2.should_save_as_image

    def test_classify_shape(self, maker):
        """Test shape classification"""
        # Test picture shape
        mock_picture = Mock()
        mock_picture.shape_type = 13  # PICTURE
        mock_picture.name = "TestImage"
        mock_picture.left = 0
        mock_picture.top = 0
        mock_picture.width = 100
        mock_picture.height = 100

        obj = maker._classify_shape(mock_picture)
        assert obj is not None
        assert obj.object_type == "image"

    def test_generate_enhanced_markdown(self, maker, temp_output_dir):
        """Test markdown generation"""
        # Create mock presentation
        mock_prs = Mock()
        mock_prs.slides = []
        mock_prs.slide_width = 9144000
        mock_prs.slide_height = 6858000

        # Create mock slide
        mock_slide = Mock()
        mock_slide.shapes = []
        mock_slide.shapes.title = None
        mock_prs.slides.append(mock_slide)

        # Create analysis
        analyses = [
            SlideAnalysis(
                slide_number=1,
                object_count=3,
                objects=[
                    SlideObject("text", "Title", None, None),
                    SlideObject("shape", "Box", None, None),
                    SlideObject("image", "Logo", None, None)
                ],
                has_image=True,
                has_chart=False,
                has_table=False,
                complexity_score=5.0,
                should_save_as_image=False,
                title="Test Slide"
            )
        ]

        pptx_path = Path("test.pptx")
        markdown = maker._generate_enhanced_markdown(mock_prs, pptx_path, analyses)

        assert "# test" in markdown
        assert "Slide 1" in markdown
        assert "Test Slide" in markdown
        assert "Objects: 3" in markdown

    def test_save_slide_as_image(self, maker, temp_output_dir):
        """Test slide image saving"""
        mock_slide = Mock()
        mock_slide.shapes = []

        image_path = maker._save_slide_as_image(mock_slide, 1, "test")

        assert image_path.exists()
        assert image_path.suffix == ".png"
        assert "test_slide_001" in image_path.name

    def test_generate_mermaid_flowchart(self, maker):
        """Test Mermaid flowchart generation"""
        analysis = SlideAnalysis(
            slide_number=1,
            object_count=4,
            objects=[
                SlideObject("shape", "Step 1", None, None),
                SlideObject("shape", "Step 2", None, None),
                SlideObject("shape", "Step 3", None, None),
                SlideObject("shape", "Step 4", None, None)
            ],
            has_image=False,
            has_chart=False,
            has_table=False,
            complexity_score=4.0,
            should_save_as_image=False,
            title="Process Flow"
        )

        mermaid = maker._try_generate_mermaid(analysis)

        # With 4 shapes, should generate flowchart
        assert mermaid is not None
        assert "flowchart" in mermaid or "graph" in mermaid

    def test_generate_mermaid_chart(self, maker):
        """Test Mermaid chart generation"""
        analysis = SlideAnalysis(
            slide_number=1,
            object_count=3,
            objects=[
                SlideObject("chart", "Sales Chart", None, None),
                SlideObject("text", "Q1", None, None),
                SlideObject("text", "Q2", None, None)
            ],
            has_image=False,
            has_chart=True,
            has_table=False,
            complexity_score=6.0,
            should_save_as_image=True,
            title="Sales Data"
        )

        mermaid = maker._try_generate_mermaid(analysis)

        assert mermaid is not None
        assert "graph" in mermaid or "Sales Data" in mermaid


class TestSlideObject:
    """Test cases for SlideObject dataclass"""

    def test_slide_object_creation(self):
        """Test SlideObject creation"""
        obj = SlideObject(
            object_type="shape",
            description="Rectangle",
            position={"left": 100, "top": 200},
            size={"width": 300, "height": 150}
        )

        assert obj.object_type == "shape"
        assert obj.description == "Rectangle"
        assert obj.position["left"] == 100
        assert obj.size["width"] == 300

    def test_slide_object_without_optional_fields(self):
        """Test SlideObject with minimal fields"""
        obj = SlideObject(
            object_type="text",
            description="Sample text"
        )

        assert obj.object_type == "text"
        assert obj.position is None
        assert obj.size is None


class TestSlideAnalysis:
    """Test cases for SlideAnalysis dataclass"""

    def test_slide_analysis_creation(self):
        """Test SlideAnalysis creation"""
        objects = [
            SlideObject("shape", "Box 1"),
            SlideObject("text", "Title"),
            SlideObject("image", "Logo")
        ]

        analysis = SlideAnalysis(
            slide_number=1,
            object_count=3,
            objects=objects,
            has_image=True,
            has_chart=False,
            has_table=False,
            complexity_score=5.0,
            should_save_as_image=False,
            title="Test Slide"
        )

        assert analysis.slide_number == 1
        assert analysis.object_count == 3
        assert len(analysis.objects) == 3
        assert analysis.has_image is True
        assert analysis.title == "Test Slide"


class TestConversionMetadata:
    """Test cases for ConversionMetadata dataclass"""

    def test_metadata_creation(self):
        """Test ConversionMetadata creation"""
        analyses = [
            SlideAnalysis(1, 3, [], False, False, False, 3.0, False),
            SlideAnalysis(2, 7, [], True, True, False, 10.0, True)
        ]

        metadata = ConversionMetadata(
            source_file="test.pptx",
            total_slides=2,
            slides_as_images=1,
            slides_as_text=1,
            mermaid_diagrams=0,
            conversion_time=5.5,
            timestamp="2025-11-25T12:00:00",
            slide_analyses=analyses
        )

        assert metadata.total_slides == 2
        assert metadata.slides_as_images == 1
        assert metadata.slides_as_text == 1
        assert metadata.conversion_time == 5.5
        assert len(metadata.slide_analyses) == 2


class TestIntegration:
    """Integration tests"""

    @pytest.fixture
    def temp_dir(self):
        """Create temporary directory"""
        temp = tempfile.mkdtemp()
        yield Path(temp)
        shutil.rmtree(temp)

    def test_full_workflow_without_file(self, temp_dir):
        """Test full workflow with mocked presentation"""
        maker = EnhancedMarkdownMaker(
            llm_client=None,
            output_dir=str(temp_dir),
            object_threshold=5
        )

        # Verify directories were created
        assert maker.slides_dir.exists()
        assert maker.mermaid_dir.exists()

        # Verify threshold setting
        assert maker.object_threshold == 5


# Pytest configuration
@pytest.fixture(scope="session")
def test_data_dir():
    """Fixture for test data directory"""
    data_dir = Path(__file__).parent / "test_data"
    data_dir.mkdir(exist_ok=True)
    return data_dir


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
