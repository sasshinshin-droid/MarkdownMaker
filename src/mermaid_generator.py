#!/usr/bin/env python3
"""
Mermaid Diagram Generator for Enhanced MarkdownMaker
Converts complex slide visuals into Mermaid diagram syntax
"""

import logging
import json
from pathlib import Path
from typing import Optional, Dict, List, Any
from enum import Enum
from dataclasses import dataclass


logger = logging.getLogger(__name__)


class DiagramType(Enum):
    """Types of diagrams that can be generated"""
    FLOWCHART = "flowchart"
    SEQUENCE = "sequenceDiagram"
    CLASS = "classDiagram"
    STATE = "stateDiagram-v2"
    ENTITY_RELATIONSHIP = "erDiagram"
    GANTT = "gantt"
    PIE = "pie"
    MINDMAP = "mindmap"
    TIMELINE = "timeline"
    UNKNOWN = "unknown"


@dataclass
class DiagramElement:
    """Represents an element in a diagram"""
    id: str
    label: str
    element_type: str
    position: Optional[Dict[str, float]] = None
    connections: Optional[List[str]] = None
    style: Optional[str] = None


class MermaidGenerator:
    """
    Generates Mermaid diagram syntax from slide analysis data.
    """

    def __init__(self, enable_styling: bool = True):
        """
        Initialize the Mermaid generator.

        Args:
            enable_styling: Enable custom styling in Mermaid output
        """
        self.enable_styling = enable_styling
        logger.info("MermaidGenerator initialized")

    def generate_from_analysis(
        self,
        analysis_data: Dict[str, Any],
        diagram_type: Optional[DiagramType] = None
    ) -> Optional[str]:
        """
        Generate Mermaid diagram from LLM analysis data.

        Args:
            analysis_data: Dictionary with slide analysis data
            diagram_type: Optional specific diagram type to generate

        Returns:
            Mermaid diagram syntax string or None
        """
        try:
            # Auto-detect diagram type if not specified
            if diagram_type is None:
                diagram_type = self._detect_diagram_type(analysis_data)

            if diagram_type == DiagramType.UNKNOWN:
                logger.warning("Could not determine diagram type")
                return None

            # Generate appropriate diagram
            generator_map = {
                DiagramType.FLOWCHART: self._generate_flowchart,
                DiagramType.SEQUENCE: self._generate_sequence_diagram,
                DiagramType.CLASS: self._generate_class_diagram,
                DiagramType.STATE: self._generate_state_diagram,
                DiagramType.PIE: self._generate_pie_chart,
                DiagramType.MINDMAP: self._generate_mindmap,
            }

            generator = generator_map.get(diagram_type)
            if generator:
                return generator(analysis_data)

            logger.warning(f"No generator available for {diagram_type}")
            return None

        except Exception as e:
            logger.error(f"Error generating Mermaid diagram: {e}")
            return None

    def _detect_diagram_type(self, analysis_data: Dict[str, Any]) -> DiagramType:
        """Detect the type of diagram from analysis data"""
        slide_type = analysis_data.get("slide_type", "").lower()
        description = analysis_data.get("description", "").lower()
        objects = analysis_data.get("objects", [])

        # Look for keywords
        if any(keyword in description for keyword in ["flow", "process", "workflow", "step"]):
            return DiagramType.FLOWCHART

        if any(keyword in description for keyword in ["sequence", "interaction", "timeline"]):
            return DiagramType.SEQUENCE

        if any(keyword in description for keyword in ["class", "object", "inheritance"]):
            return DiagramType.CLASS

        if any(keyword in description for keyword in ["state", "status", "transition"]):
            return DiagramType.STATE

        if any(keyword in description for keyword in ["percentage", "proportion", "distribution"]):
            return DiagramType.PIE

        if any(keyword in description for keyword in ["hierarchy", "tree", "mind map"]):
            return DiagramType.MINDMAP

        # Count shapes to determine if it's likely a flowchart
        shape_count = sum(1 for obj in objects if obj.get("type") == "shape")
        if shape_count >= 3:
            return DiagramType.FLOWCHART

        return DiagramType.UNKNOWN

    def _generate_flowchart(self, analysis_data: Dict[str, Any]) -> str:
        """Generate a flowchart diagram"""
        mermaid = ["flowchart TD"]

        objects = analysis_data.get("objects", [])
        title = analysis_data.get("title", "Process")

        # Extract nodes
        nodes = []
        for i, obj in enumerate(objects):
            if obj.get("type") in ["shape", "text"]:
                node_id = f"node{i+1}"
                label = self._clean_label(obj.get("description", f"Step {i+1}"))

                # Determine node shape based on description
                if any(keyword in label.lower() for keyword in ["start", "begin"]):
                    node_def = f'    {node_id}(["{label}"])'
                elif any(keyword in label.lower() for keyword in ["end", "finish"]):
                    node_def = f'    {node_id}(["{label}"])'
                elif any(keyword in label.lower() for keyword in ["decision", "if", "?"]):
                    node_def = f'    {node_id}{{{label}}}'
                else:
                    node_def = f'    {node_id}["{label}"]'

                nodes.append(node_def)
                mermaid.append(node_def)

        # Add connections (simple linear for now)
        for i in range(len(nodes) - 1):
            mermaid.append(f"    node{i+1} --> node{i+2}")

        # Add styling if enabled
        if self.enable_styling:
            mermaid.append("    style node1 fill:#e1f5e1")
            if len(nodes) > 1:
                mermaid.append(f"    style node{len(nodes)} fill:#ffe1e1")

        return "\n".join(mermaid)

    def _generate_sequence_diagram(self, analysis_data: Dict[str, Any]) -> str:
        """Generate a sequence diagram"""
        mermaid = ["sequenceDiagram"]

        objects = analysis_data.get("objects", [])

        # Extract participants
        participants = []
        for i, obj in enumerate(objects[:5]):  # Limit to 5 participants
            label = self._clean_label(obj.get("description", f"Actor{i+1}"))
            participant_id = f"P{i+1}"
            participants.append(participant_id)
            mermaid.append(f"    participant {participant_id} as {label}")

        # Add interactions
        if len(participants) >= 2:
            for i in range(len(participants) - 1):
                mermaid.append(f"    {participants[i]}->>+{participants[i+1]}: Request")
                mermaid.append(f"    {participants[i+1]}-->>-{participants[i]}: Response")

        return "\n".join(mermaid)

    def _generate_class_diagram(self, analysis_data: Dict[str, Any]) -> str:
        """Generate a class diagram"""
        mermaid = ["classDiagram"]

        objects = analysis_data.get("objects", [])

        # Create classes from objects
        for i, obj in enumerate(objects[:5]):  # Limit to 5 classes
            class_name = self._clean_label(obj.get("description", f"Class{i+1}"))
            class_name = class_name.replace(" ", "")[:20]

            mermaid.append(f"    class {class_name} {{")
            mermaid.append(f"        +attribute{i+1}")
            mermaid.append(f"        +method{i+1}()")
            mermaid.append("    }")

        # Add relationships if multiple classes
        if len(objects) > 1:
            for i in range(min(len(objects) - 1, 4)):
                class1 = self._clean_label(objects[i].get("description", f"Class{i+1}")).replace(" ", "")[:20]
                class2 = self._clean_label(objects[i+1].get("description", f"Class{i+2}")).replace(" ", "")[:20]
                mermaid.append(f"    {class1} --> {class2}")

        return "\n".join(mermaid)

    def _generate_state_diagram(self, analysis_data: Dict[str, Any]) -> str:
        """Generate a state diagram"""
        mermaid = ["stateDiagram-v2"]

        objects = analysis_data.get("objects", [])

        # Add states
        mermaid.append("    [*] --> State1")

        for i, obj in enumerate(objects[:5], 1):
            state_name = f"State{i}"
            label = self._clean_label(obj.get("description", f"State {i}"))

            mermaid.append(f"    {state_name}: {label}")

            if i < len(objects[:5]):
                mermaid.append(f"    {state_name} --> State{i+1}")

        # Final state
        final_state = min(len(objects[:5]), 5)
        mermaid.append(f"    State{final_state} --> [*]")

        return "\n".join(mermaid)

    def _generate_pie_chart(self, analysis_data: Dict[str, Any]) -> str:
        """Generate a pie chart"""
        mermaid = ['pie title Slide Content Distribution']

        objects = analysis_data.get("objects", [])

        # Count object types
        type_counts = {}
        for obj in objects:
            obj_type = obj.get("type", "other")
            type_counts[obj_type] = type_counts.get(obj_type, 0) + 1

        # Add to pie chart
        for obj_type, count in type_counts.items():
            label = obj_type.capitalize()
            mermaid.append(f'    "{label}": {count}')

        return "\n".join(mermaid)

    def _generate_mindmap(self, analysis_data: Dict[str, Any]) -> str:
        """Generate a mindmap"""
        mermaid = ["mindmap"]

        title = analysis_data.get("title", "Main Topic")
        mermaid.append(f"  root(({title}))")

        objects = analysis_data.get("objects", [])

        # Add branches
        for i, obj in enumerate(objects[:6]):
            label = self._clean_label(obj.get("description", f"Branch {i+1}"))
            mermaid.append(f"    {label}")

        return "\n".join(mermaid)

    def _clean_label(self, label: str) -> str:
        """Clean and format label for Mermaid"""
        # Remove quotes and special characters
        label = label.replace('"', "'").replace('\n', ' ')
        # Truncate if too long
        if len(label) > 50:
            label = label[:47] + "..."
        return label

    def generate_from_objects(
        self,
        objects: List[Dict[str, Any]],
        diagram_type: DiagramType = DiagramType.FLOWCHART
    ) -> str:
        """
        Generate Mermaid diagram directly from object list.

        Args:
            objects: List of object dictionaries
            diagram_type: Type of diagram to generate

        Returns:
            Mermaid diagram syntax
        """
        analysis_data = {
            "objects": objects,
            "slide_type": diagram_type.value,
            "description": "",
            "title": "Generated Diagram"
        }

        return self.generate_from_analysis(analysis_data, diagram_type)

    def validate_mermaid_syntax(self, mermaid_code: str) -> bool:
        """
        Validate Mermaid syntax (basic validation).

        Args:
            mermaid_code: Mermaid diagram code

        Returns:
            True if syntax appears valid
        """
        if not mermaid_code or not mermaid_code.strip():
            return False

        lines = mermaid_code.strip().split('\n')
        first_line = lines[0].strip()

        # Check if starts with valid diagram type
        valid_starts = [
            "flowchart", "graph", "sequenceDiagram", "classDiagram",
            "stateDiagram", "erDiagram", "gantt", "pie", "mindmap", "timeline"
        ]

        return any(first_line.startswith(start) for start in valid_starts)

    def save_mermaid_file(
        self,
        mermaid_code: str,
        output_path: Path,
        include_metadata: bool = True
    ) -> Path:
        """
        Save Mermaid diagram to file.

        Args:
            mermaid_code: Mermaid diagram code
            output_path: Output file path
            include_metadata: Include metadata comments

        Returns:
            Path to saved file
        """
        output_path.parent.mkdir(parents=True, exist_ok=True)

        content = []

        if include_metadata:
            content.append("%%{init: {'theme':'base'}}%%")
            content.append(f"%% Generated by Enhanced MarkdownMaker")
            content.append(f"%% Generated at: {Path(output_path).name}")
            content.append("")

        content.append(mermaid_code)

        output_path.write_text("\n".join(content), encoding='utf-8')
        logger.info(f"Saved Mermaid diagram to {output_path}")

        return output_path

    def batch_generate(
        self,
        analyses: List[Dict[str, Any]],
        output_dir: Path
    ) -> Dict[int, Path]:
        """
        Generate Mermaid diagrams for multiple slides in batch.

        Args:
            analyses: List of slide analysis dictionaries
            output_dir: Output directory for Mermaid files

        Returns:
            Dictionary mapping slide numbers to file paths
        """
        results = {}

        for i, analysis in enumerate(analyses, 1):
            try:
                mermaid_code = self.generate_from_analysis(analysis)

                if mermaid_code and self.validate_mermaid_syntax(mermaid_code):
                    output_path = output_dir / f"slide_{i:03d}.mmd"
                    self.save_mermaid_file(mermaid_code, output_path)
                    results[i] = output_path
                else:
                    logger.warning(f"Could not generate valid Mermaid for slide {i}")
                    results[i] = None

            except Exception as e:
                logger.error(f"Error generating Mermaid for slide {i}: {e}")
                results[i] = None

        logger.info(f"Batch generated {len([r for r in results.values() if r])} Mermaid diagrams")
        return results


def create_sample_diagrams(output_dir: Path):
    """Create sample Mermaid diagrams for testing"""
    generator = MermaidGenerator()

    samples = {
        "flowchart_sample.mmd": {
            "objects": [
                {"type": "shape", "description": "Start Process"},
                {"type": "shape", "description": "Check Input"},
                {"type": "shape", "description": "Process Data"},
                {"type": "shape", "description": "Generate Output"},
                {"type": "shape", "description": "End Process"}
            ],
            "slide_type": "flowchart",
            "description": "A simple process flow"
        },
        "sequence_sample.mmd": {
            "objects": [
                {"type": "text", "description": "User"},
                {"type": "text", "description": "API"},
                {"type": "text", "description": "Database"}
            ],
            "slide_type": "sequence",
            "description": "API interaction sequence"
        },
        "pie_sample.mmd": {
            "objects": [
                {"type": "shape", "description": "Category A"},
                {"type": "shape", "description": "Category B"},
                {"type": "chart", "description": "Category C"}
            ],
            "slide_type": "pie",
            "description": "Data distribution"
        }
    }

    output_dir.mkdir(parents=True, exist_ok=True)

    for filename, sample_data in samples.items():
        mermaid_code = generator.generate_from_analysis(sample_data)
        if mermaid_code:
            output_path = output_dir / filename
            generator.save_mermaid_file(mermaid_code, output_path)
            print(f"Created: {output_path}")


if __name__ == "__main__":
    # Test the Mermaid generator
    output_dir = Path("output/mermaid_samples")
    create_sample_diagrams(output_dir)
    print(f"\nSample Mermaid diagrams created in {output_dir}")
