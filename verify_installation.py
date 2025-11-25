#!/usr/bin/env python3
"""
Installation verification script for Enhanced MarkdownMaker
Checks all dependencies and system components
"""

import sys
from pathlib import Path

print("=" * 70)
print("Enhanced MarkdownMaker - Installation Verification")
print("=" * 70)

# Check Python version
print("\n[1] Python Version")
version = sys.version_info
print(f"   Python {version.major}.{version.minor}.{version.micro}")
if version.major >= 3 and version.minor >= 8:
    print("   ✓ Python version OK (>= 3.8)")
else:
    print("   ✗ Python version too old (need >= 3.8)")
    sys.exit(1)

# Check core dependencies
print("\n[2] Core Dependencies")
dependencies = {
    'markitdown': 'Microsoft MarkItDown library',
    'pptx': 'python-pptx for PowerPoint parsing',
    'PIL': 'Pillow for image processing',
    'yaml': 'PyYAML for configuration',
}

missing = []
for module, description in dependencies.items():
    try:
        __import__(module)
        print(f"   ✓ {module:15s} - {description}")
    except ImportError:
        print(f"   ✗ {module:15s} - {description} [MISSING]")
        missing.append(module)

if missing:
    print(f"\n   Missing dependencies: {', '.join(missing)}")
    print("   Install with: pip install -r requirements.txt")

# Check optional dependencies
print("\n[3] Optional Dependencies")
optional = {
    'openai': 'OpenAI for LLM integration',
    'pytest': 'Testing framework',
    'black': 'Code formatter',
}

for module, description in optional.items():
    try:
        __import__(module)
        print(f"   ✓ {module:15s} - {description}")
    except ImportError:
        print(f"   - {module:15s} - {description} [Not installed]")

# Check project structure
print("\n[4] Project Structure")
required_files = [
    'src/enhanced_markdown_maker.py',
    'src/image_processor.py',
    'src/mermaid_generator.py',
    'src/config.py',
    'src/cli.py',
    'src/__init__.py',
    'requirements.txt',
    'setup.py',
    'README.md',
]

project_root = Path(__file__).parent
for file_path in required_files:
    full_path = project_root / file_path
    if full_path.exists():
        print(f"   ✓ {file_path}")
    else:
        print(f"   ✗ {file_path} [MISSING]")

# Check if modules can be imported
print("\n[5] Module Imports")
sys.path.insert(0, str(project_root / "src"))

modules_to_test = [
    ('enhanced_markdown_maker', 'EnhancedMarkdownMaker'),
    ('image_processor', 'ImageProcessor'),
    ('mermaid_generator', 'MermaidGenerator'),
    ('config', 'ConfigManager'),
]

import_errors = []
for module_name, class_name in modules_to_test:
    try:
        module = __import__(module_name)
        if hasattr(module, class_name):
            print(f"   ✓ {module_name:25s} - {class_name} available")
        else:
            print(f"   ✗ {module_name:25s} - {class_name} not found")
            import_errors.append(module_name)
    except Exception as e:
        print(f"   ✗ {module_name:25s} - Import error: {e}")
        import_errors.append(module_name)

# Check output directories
print("\n[6] Output Directories")
output_dir = project_root / "output"
if output_dir.exists():
    print(f"   ✓ output/ directory exists")
else:
    print(f"   - output/ directory will be created on first run")

# System tools
print("\n[7] System Tools (Optional)")
tools = {
    'libreoffice': 'For slide image extraction',
    'unoconv': 'Alternative slide extraction',
}

import shutil
for tool, description in tools.items():
    if shutil.which(tool):
        print(f"   ✓ {tool:15s} - {description}")
    else:
        print(f"   - {tool:15s} - {description} [Not found]")

# Summary
print("\n" + "=" * 70)
print("Summary")
print("=" * 70)

if missing:
    print(f"\n✗ Installation INCOMPLETE")
    print(f"  Missing required dependencies: {', '.join(missing)}")
    print(f"  Run: pip install -r requirements.txt")
    sys.exit(1)
elif import_errors:
    print(f"\n⚠ Installation PARTIAL")
    print(f"  Some modules have import errors: {', '.join(import_errors)}")
    print(f"  Check error messages above")
    sys.exit(1)
else:
    print(f"\n✓ Installation COMPLETE")
    print(f"\n  All required components are installed and working!")
    print(f"\n  Next steps:")
    print(f"    1. Run demo: python demo.py")
    print(f"    2. Try examples: python examples/basic_usage.py")
    print(f"    3. Run tests: pytest tests/")
    print(f"    4. See documentation: README.md")
    print(f"    5. Quick start: QUICKSTART.md")

print()
