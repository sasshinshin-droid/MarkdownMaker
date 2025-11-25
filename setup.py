#!/usr/bin/env python3
"""
Setup script for Enhanced MarkdownMaker
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README for long description
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding="utf-8") if readme_file.exists() else ""

# Read requirements
requirements_file = Path(__file__).parent / "requirements.txt"
requirements = []
if requirements_file.exists():
    with open(requirements_file) as f:
        requirements = [
            line.strip()
            for line in f
            if line.strip() and not line.startswith('#')
        ]

# Core requirements (required for basic functionality)
core_requirements = [
    "markitdown>=0.1.3",
    "python-pptx>=0.6.21",
    "Pillow>=9.0.0",
    "pyyaml>=6.0.0",
    "click>=8.0.0",
]

# Optional requirements for advanced features
extras_require = {
    'llm': [
        'openai>=1.0.0',
    ],
    'dev': [
        'pytest>=7.0.0',
        'pytest-cov>=4.0.0',
        'pytest-asyncio>=0.21.0',
        'black>=23.0.0',
        'flake8>=6.0.0',
        'mypy>=1.0.0',
        'types-PyYAML>=6.0.0',
        'types-Pillow>=9.0.0',
    ],
    'docs': [
        'sphinx>=5.0.0',
        'sphinx-rtd-theme>=1.0.0',
    ],
    'ui': [
        'tqdm>=4.65.0',
        'colorama>=0.4.6',
        'rich>=13.0.0',
    ]
}

# All extras
extras_require['all'] = list(set(sum(extras_require.values(), [])))

setup(
    name="enhanced-markdown-maker",
    version="1.0.0",
    author="Enhanced MarkdownMaker Contributors",
    author_email="contact@example.com",
    description="PowerPoint to Markdown converter with enhanced object counting and Mermaid diagram generation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/enhanced-markdown-maker",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/enhanced-markdown-maker/issues",
        "Source": "https://github.com/yourusername/enhanced-markdown-maker",
        "Documentation": "https://enhanced-markdown-maker.readthedocs.io",
    },
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Text Processing :: Markup :: Markdown",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
    install_requires=core_requirements,
    extras_require=extras_require,
    entry_points={
        "console_scripts": [
            "markdown-maker=cli:main",
            "mm=cli:main",
        ],
    },
    include_package_data=True,
    keywords="powerpoint markdown converter pptx markitdown mermaid",
    zip_safe=False,
)
