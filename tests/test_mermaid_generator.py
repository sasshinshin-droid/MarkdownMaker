#!/usr/bin/env python3
"""
Test suite for Mermaid Generator
Tests diagram generation functionality
"""

import pytest
import sys
from pathlib import Path
import tempfile
import shutil

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from mermaid_generator import (
    MermaidGenerator,
    DiagramType,
    DiagramElement
)


class TestMermaidGenerator:
    """Test cases for MermaidGenerator class"""

    @pytest.fixture
    def generator(self):
        """Create MermaidGenerator instance"""
        return MermaidGenerator(enable_styling=True)

    @pytest.fixture
    def temp_dir(self):
        """Create temporary directory"""
        temp = tempfile.mkdtemp()
        yield Path(temp)
        shutil.rmtree(temp)

    def test_initialization(self, generator):
        """Test MermaidGenerator initialization"""
        assert generator.enable_styling is True

    def test_detect_flowchart_type(self, generator):
        """Test flowchart detection"""
        analysis = {
            "slide_type": "diagram",
            "description": "This shows a workflow process with multiple steps",
            "objects": [
                {"type": "shape", "description": "Step 1"},
                {"type": "shape", "description": "Step 2"},
                {"type": "shape", "description": "Step 3"}
            ]
        }

        diagram_type = generator._detect_diagram_type(analysis)
        assert diagram_type == DiagramType.FLOWCHART

    def test_detect_sequence_type(self, generator):
        """Test sequence diagram detection"""
        analysis = {
            "slide_type": "interaction",
            "description": "API interaction sequence showing request and response",
            "objects": [
                {"type": "text", "description": "Client"},
                {"type": "text", "description": "Server"}
            ]
        }

        diagram_type = generator._detect_diagram_type(analysis)
        assert diagram_type == DiagramType.SEQUENCE

    def test_detect_pie_type(self, generator):
        """Test pie chart detection"""
        analysis = {
            "slide_type": "chart",
            "description": "Market share distribution by percentage",
            "objects": [
                {"type": "chart", "description": "Pie Chart"}
            ]
        }

        diagram_type = generator._detect_diagram_type(analysis)
        assert diagram_type == DiagramType.PIE

    def test_generate_flowchart(self, generator):
        """Test flowchart generation"""
        analysis = {
            "title": "Process Flow",
            "objects": [
                {"type": "shape", "description": "Start Process"},
                {"type": "shape", "description": "Validate Input"},
                {"type": "shape", "description": "Process Data"},
                {"type": "shape", "description": "End Process"}
            ],
            "slide_type": "flowchart"
        }

        mermaid = generator._generate_flowchart(analysis)

        assert "flowchart TD" in mermaid
        assert "Start Process" in mermaid
        assert "End Process" in mermaid
        assert "-->" in mermaid

    def test_generate_sequence_diagram(self, generator):
        """Test sequence diagram generation"""
        analysis = {
            "title": "API Flow",
            "objects": [
                {"type": "text", "description": "User"},
                {"type": "text", "description": "API"},
                {"type": "text", "description": "Database"}
            ],
            "slide_type": "sequence"
        }

        mermaid = generator._generate_sequence_diagram(analysis)

        assert "sequenceDiagram" in mermaid
        assert "participant" in mermaid
        assert "User" in mermaid or "P1" in mermaid

    def test_generate_class_diagram(self, generator):
        """Test class diagram generation"""
        analysis = {
            "title": "Class Structure",
            "objects": [
                {"type": "shape", "description": "UserClass"},
                {"type": "shape", "description": "OrderClass"}
            ],
            "slide_type": "class"
        }

        mermaid = generator._generate_class_diagram(analysis)

        assert "classDiagram" in mermaid
        assert "class" in mermaid
        assert "-->" in mermaid or "UserClass" in mermaid

    def test_generate_state_diagram(self, generator):
        """Test state diagram generation"""
        analysis = {
            "title": "State Machine",
            "objects": [
                {"type": "shape", "description": "Initial State"},
                {"type": "shape", "description": "Processing State"},
                {"type": "shape", "description": "Final State"}
            ],
            "slide_type": "state"
        }

        mermaid = generator._generate_state_diagram(analysis)

        assert "stateDiagram-v2" in mermaid
        assert "[*]" in mermaid

    def test_generate_pie_chart(self, generator):
        """Test pie chart generation"""
        analysis = {
            "title": "Distribution",
            "objects": [
                {"type": "shape", "description": "Category A"},
                {"type": "shape", "description": "Category B"},
                {"type": "chart", "description": "Category C"}
            ],
            "slide_type": "pie"
        }

        mermaid = generator._generate_pie_chart(analysis)

        assert "pie" in mermaid
        assert "title" in mermaid

    def test_generate_mindmap(self, generator):
        """Test mindmap generation"""
        analysis = {
            "title": "Main Topic",
            "objects": [
                {"type": "text", "description": "Subtopic 1"},
                {"type": "text", "description": "Subtopic 2"},
                {"type": "text", "description": "Subtopic 3"}
            ],
            "slide_type": "mindmap"
        }

        mermaid = generator._generate_mindmap(analysis)

        assert "mindmap" in mermaid
        assert "Main Topic" in mermaid

    def test_clean_label(self, generator):
        """Test label cleaning"""
        # Test quote replacement
        label1 = 'Text with "quotes"'
        cleaned1 = generator._clean_label(label1)
        assert '"' not in cleaned1
        assert "'" in cleaned1

        # Test newline removal
        label2 = "Text\nwith\nnewlines"
        cleaned2 = generator._clean_label(label2)
        assert "\n" not in cleaned2

        # Test truncation
        label3 = "A" * 100
        cleaned3 = generator._clean_label(label3)
        assert len(cleaned3) <= 50
        assert cleaned3.endswith("...")

    def test_validate_mermaid_syntax(self, generator):
        """Test Mermaid syntax validation"""
        # Valid flowchart
        valid1 = "flowchart TD\n    A --> B"
        assert generator.validate_mermaid_syntax(valid1) is True

        # Valid sequence
        valid2 = "sequenceDiagram\n    A->>B: Message"
        assert generator.validate_mermaid_syntax(valid2) is True

        # Invalid (empty)
        invalid1 = ""
        assert generator.validate_mermaid_syntax(invalid1) is False

        # Invalid (wrong start)
        invalid2 = "invalid diagram\n    A --> B"
        assert generator.validate_mermaid_syntax(invalid2) is False

    def test_save_mermaid_file(self, generator, temp_dir):
        """Test saving Mermaid to file"""
        mermaid_code = "flowchart TD\n    A[Start] --> B[End]"
        output_path = temp_dir / "test.mmd"

        result_path = generator.save_mermaid_file(
            mermaid_code,
            output_path,
            include_metadata=True
        )

        assert result_path.exists()
        content = result_path.read_text()
        assert "flowchart TD" in content
        assert "Generated by Enhanced MarkdownMaker" in content

    def test_save_mermaid_file_without_metadata(self, generator, temp_dir):
        """Test saving Mermaid without metadata"""
        mermaid_code = "flowchart TD\n    A --> B"
        output_path = temp_dir / "test_no_meta.mmd"

        result_path = generator.save_mermaid_file(
            mermaid_code,
            output_path,
            include_metadata=False
        )

        content = result_path.read_text()
        assert "Generated by" not in content
        assert "flowchart TD" in content

    def test_batch_generate(self, generator, temp_dir):
        """Test batch generation"""
        analyses = [
            {
                "title": "Flow 1",
                "objects": [
                    {"type": "shape", "description": "Step A"},
                    {"type": "shape", "description": "Step B"}
                ],
                "slide_type": "flowchart",
                "description": "Process flow"
            },
            {
                "title": "Sequence 1",
                "objects": [
                    {"type": "text", "description": "User"},
                    {"type": "text", "description": "System"}
                ],
                "slide_type": "sequence",
                "description": "User interaction"
            }
        ]

        results = generator.batch_generate(analyses, temp_dir)

        assert len(results) == 2
        assert 1 in results
        assert 2 in results

        # Check that files were created
        if results[1]:
            assert results[1].exists()

    def test_generate_from_objects(self, generator):
        """Test direct generation from objects"""
        objects = [
            {"type": "shape", "description": "Node 1"},
            {"type": "shape", "description": "Node 2"},
            {"type": "shape", "description": "Node 3"}
        ]

        mermaid = generator.generate_from_objects(
            objects,
            DiagramType.FLOWCHART
        )

        assert mermaid is not None
        assert "flowchart" in mermaid or "graph" in mermaid


class TestDiagramElement:
    """Test cases for DiagramElement dataclass"""

    def test_diagram_element_creation(self):
        """Test DiagramElement creation"""
        element = DiagramElement(
            id="node1",
            label="Start",
            element_type="start",
            position={"x": 100, "y": 200},
            connections=["node2", "node3"],
            style="fill:#f9f"
        )

        assert element.id == "node1"
        assert element.label == "Start"
        assert element.element_type == "start"
        assert len(element.connections) == 2

    def test_diagram_element_minimal(self):
        """Test DiagramElement with minimal fields"""
        element = DiagramElement(
            id="node1",
            label="Node",
            element_type="process"
        )

        assert element.position is None
        assert element.connections is None
        assert element.style is None


class TestDiagramType:
    """Test cases for DiagramType enum"""

    def test_diagram_type_values(self):
        """Test DiagramType enum values"""
        assert DiagramType.FLOWCHART.value == "flowchart"
        assert DiagramType.SEQUENCE.value == "sequenceDiagram"
        assert DiagramType.CLASS.value == "classDiagram"
        assert DiagramType.STATE.value == "stateDiagram-v2"
        assert DiagramType.PIE.value == "pie"
        assert DiagramType.MINDMAP.value == "mindmap"
        assert DiagramType.UNKNOWN.value == "unknown"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
