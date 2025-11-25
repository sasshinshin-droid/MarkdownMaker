# Enhanced MarkdownMaker - Quick Start Guide

Get started with Enhanced MarkdownMaker in 5 minutes!

## Installation

```bash
# 1. Navigate to the project directory
cd markdown_maker_system

# 2. Install dependencies
pip install -r requirements.txt

# 3. (Optional) Install in development mode
pip install -e .
```

## Your First Conversion

### Method 1: Using the CLI

```bash
# Convert a PowerPoint file
python src/cli.py convert your_presentation.pptx

# Output will be in: output/your_presentation.md
```

### Method 2: Using Python API

```python
from src.enhanced_markdown_maker import EnhancedMarkdownMaker

# Initialize converter
maker = EnhancedMarkdownMaker(
    output_dir="output",
    object_threshold=5
)

# Convert
markdown, metadata = maker.convert_powerpoint("presentation.pptx")

print(f"✓ Converted {metadata.total_slides} slides!")
```

## Common Use Cases

### 1. Simple Conversion

```bash
python src/cli.py convert presentation.pptx
```

### 2. Lower Threshold (More Images)

```bash
python src/cli.py convert presentation.pptx --threshold 3
```

### 3. Enable Mermaid Diagrams

```bash
python src/cli.py convert presentation.pptx --enable-mermaid
```

### 4. Custom Output Directory

```bash
python src/cli.py convert presentation.pptx --output-dir ./docs
```

### 5. Batch Convert Multiple Files

```bash
python src/cli.py batch ./presentations/ --output-dir ./output
```

### 6. Analyze Without Converting

```bash
python src/cli.py analyze presentation.pptx --verbose
```

## Configuration

Create `markdown_maker.yaml` in your project directory:

```yaml
conversion:
  object_threshold: 5
  output_dir: "output"
  enable_mermaid: true

image:
  width: 1920
  height: 1080
  dpi: 300

llm:
  enabled: false  # Set to true if using OpenAI
  model: "gpt-4o"
```

## With LLM Support (Optional)

```bash
# 1. Install OpenAI
pip install openai

# 2. Set API key
export OPENAI_API_KEY=your-api-key

# 3. Use in code
from openai import OpenAI
from src.enhanced_markdown_maker import EnhancedMarkdownMaker

client = OpenAI()
maker = EnhancedMarkdownMaker(llm_client=client)
```

## Running the Demo

```bash
# See all features in action
python demo.py
```

## Running Tests

```bash
# Run all tests
pytest tests/

# Run with coverage
pytest --cov=src tests/
```

## Directory Structure After First Run

```
output/
├── presentation.md           # Your converted Markdown
├── slides/                   # Complex slide images
│   └── presentation_slide_001.png
├── mermaid/                  # Mermaid diagrams
│   └── slide_001.mmd
└── metadata.json            # Conversion statistics
```

## Troubleshooting

### Issue: "Module not found"
```bash
# Make sure you're in the right directory
cd markdown_maker_system

# Install dependencies again
pip install -r requirements.txt
```

### Issue: "Can't find presentation file"
```bash
# Use absolute path
python src/cli.py convert /full/path/to/presentation.pptx
```

### Issue: Images not extracting
- Install LibreOffice for full image support
- System will use placeholders if LibreOffice unavailable

## Next Steps

1. Read the full [README.md](README.md) for detailed documentation
2. Explore [examples/basic_usage.py](examples/basic_usage.py) for more examples
3. Check out the test suite in `tests/` for advanced usage
4. Customize configuration in `markdown_maker.yaml`

## Getting Help

- Read documentation: [README.md](README.md)
- Check examples: `examples/` directory
- Run demo: `python demo.py`
- View CLI help: `python src/cli.py --help`

## Quick Reference

### CLI Commands
```bash
convert     # Convert single file
batch       # Convert multiple files
analyze     # Analyze without converting
mermaid     # Generate Mermaid diagram
```

### Common Options
```bash
-o, --output-dir     # Output directory
-t, --threshold      # Object count threshold
-m, --enable-mermaid # Enable Mermaid generation
-v, --verbose        # Verbose output
```

### Python API Quick Reference
```python
# Import
from src.enhanced_markdown_maker import EnhancedMarkdownMaker

# Initialize
maker = EnhancedMarkdownMaker(output_dir="output")

# Convert
markdown, metadata = maker.convert_powerpoint("file.pptx")

# Access results
print(metadata.total_slides)
print(metadata.slides_as_images)
```

---

**Happy Converting!** 🎉

For more details, see [README.md](README.md)
