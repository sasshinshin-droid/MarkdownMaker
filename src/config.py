#!/usr/bin/env python3
"""
Configuration management for Enhanced MarkdownMaker
Supports YAML, JSON, and environment variable configuration
"""

import os
import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict, field

try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False
    logging.warning("PyYAML not installed. YAML config support disabled.")


logger = logging.getLogger(__name__)


@dataclass
class ImageConfig:
    """Image processing configuration"""
    width: int = 1920
    height: int = 1080
    dpi: int = 300
    quality: int = 95
    thumbnail_size: tuple = (400, 300)


@dataclass
class MermaidConfig:
    """Mermaid diagram generation configuration"""
    enabled: bool = True
    enable_styling: bool = True
    max_nodes: int = 10
    auto_detect: bool = True


@dataclass
class ConversionConfig:
    """Main conversion configuration"""
    object_threshold: int = 5
    output_dir: str = "output"
    enable_mermaid: bool = True
    save_metadata: bool = True
    create_toc: bool = True
    watermark_images: bool = False
    watermark_text: str = "Enhanced MarkdownMaker"


@dataclass
class LLMConfig:
    """LLM integration configuration"""
    enabled: bool = False
    model: str = "gpt-4o"
    api_key: Optional[str] = None
    base_url: Optional[str] = None
    temperature: float = 0.3
    max_tokens: int = 2000


@dataclass
class MarkdownMakerConfig:
    """Complete configuration for Enhanced MarkdownMaker"""
    conversion: ConversionConfig = field(default_factory=ConversionConfig)
    image: ImageConfig = field(default_factory=ImageConfig)
    mermaid: MermaidConfig = field(default_factory=MermaidConfig)
    llm: LLMConfig = field(default_factory=LLMConfig)


class ConfigManager:
    """Manages configuration loading and saving"""

    DEFAULT_CONFIG_LOCATIONS = [
        "markdown_maker.yaml",
        "markdown_maker.yml",
        "markdown_maker.json",
        "~/.markdown_maker/config.yaml",
        "~/.markdown_maker/config.json",
    ]

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize ConfigManager.

        Args:
            config_path: Optional path to config file
        """
        self.config_path = config_path
        self.config = self._load_config()

    def _load_config(self) -> MarkdownMakerConfig:
        """Load configuration from file or defaults"""
        # Try to load from specified path
        if self.config_path:
            return self._load_from_file(Path(self.config_path))

        # Try default locations
        for location in self.DEFAULT_CONFIG_LOCATIONS:
            path = Path(location).expanduser()
            if path.exists():
                logger.info(f"Loading config from: {path}")
                return self._load_from_file(path)

        # Load from environment variables
        config = self._load_from_env()
        if config:
            logger.info("Loaded config from environment variables")
            return config

        # Return defaults
        logger.info("Using default configuration")
        return MarkdownMakerConfig()

    def _load_from_file(self, path: Path) -> MarkdownMakerConfig:
        """Load configuration from file"""
        try:
            with open(path, 'r') as f:
                if path.suffix in ['.yaml', '.yml']:
                    if not YAML_AVAILABLE:
                        raise RuntimeError("PyYAML not installed")
                    data = yaml.safe_load(f)
                elif path.suffix == '.json':
                    data = json.load(f)
                else:
                    raise ValueError(f"Unsupported config file format: {path.suffix}")

            return self._dict_to_config(data)

        except Exception as e:
            logger.error(f"Error loading config from {path}: {e}")
            return MarkdownMakerConfig()

    def _load_from_env(self) -> Optional[MarkdownMakerConfig]:
        """Load configuration from environment variables"""
        env_config = {}

        # Conversion settings
        if os.getenv('MM_OBJECT_THRESHOLD'):
            env_config.setdefault('conversion', {})['object_threshold'] = int(
                os.getenv('MM_OBJECT_THRESHOLD')
            )

        if os.getenv('MM_OUTPUT_DIR'):
            env_config.setdefault('conversion', {})['output_dir'] = os.getenv('MM_OUTPUT_DIR')

        if os.getenv('MM_ENABLE_MERMAID'):
            env_config.setdefault('conversion', {})['enable_mermaid'] = (
                os.getenv('MM_ENABLE_MERMAID').lower() == 'true'
            )

        # Image settings
        if os.getenv('MM_IMAGE_WIDTH'):
            env_config.setdefault('image', {})['width'] = int(os.getenv('MM_IMAGE_WIDTH'))

        if os.getenv('MM_IMAGE_HEIGHT'):
            env_config.setdefault('image', {})['height'] = int(os.getenv('MM_IMAGE_HEIGHT'))

        if os.getenv('MM_IMAGE_DPI'):
            env_config.setdefault('image', {})['dpi'] = int(os.getenv('MM_IMAGE_DPI'))

        # LLM settings
        if os.getenv('MM_LLM_ENABLED'):
            env_config.setdefault('llm', {})['enabled'] = (
                os.getenv('MM_LLM_ENABLED').lower() == 'true'
            )

        if os.getenv('MM_LLM_MODEL'):
            env_config.setdefault('llm', {})['model'] = os.getenv('MM_LLM_MODEL')

        if os.getenv('MM_LLM_API_KEY') or os.getenv('OPENAI_API_KEY'):
            env_config.setdefault('llm', {})['api_key'] = (
                os.getenv('MM_LLM_API_KEY') or os.getenv('OPENAI_API_KEY')
            )

        if os.getenv('MM_LLM_BASE_URL'):
            env_config.setdefault('llm', {})['base_url'] = os.getenv('MM_LLM_BASE_URL')

        if not env_config:
            return None

        return self._dict_to_config(env_config)

    def _dict_to_config(self, data: Dict[str, Any]) -> MarkdownMakerConfig:
        """Convert dictionary to MarkdownMakerConfig"""
        try:
            # Parse nested configs
            conversion_data = data.get('conversion', {})
            image_data = data.get('image', {})
            mermaid_data = data.get('mermaid', {})
            llm_data = data.get('llm', {})

            conversion = ConversionConfig(**conversion_data)
            image = ImageConfig(**image_data)
            mermaid = MermaidConfig(**mermaid_data)
            llm = LLMConfig(**llm_data)

            return MarkdownMakerConfig(
                conversion=conversion,
                image=image,
                mermaid=mermaid,
                llm=llm
            )

        except Exception as e:
            logger.error(f"Error parsing config: {e}")
            return MarkdownMakerConfig()

    def save_config(self, path: Path, format: str = 'yaml'):
        """
        Save current configuration to file.

        Args:
            path: Output file path
            format: Output format ('yaml' or 'json')
        """
        config_dict = asdict(self.config)

        try:
            with open(path, 'w') as f:
                if format == 'yaml':
                    if not YAML_AVAILABLE:
                        raise RuntimeError("PyYAML not installed")
                    yaml.dump(config_dict, f, default_flow_style=False, indent=2)
                elif format == 'json':
                    json.dump(config_dict, f, indent=2)
                else:
                    raise ValueError(f"Unsupported format: {format}")

            logger.info(f"Saved config to: {path}")

        except Exception as e:
            logger.error(f"Error saving config to {path}: {e}")
            raise

    def get_conversion_config(self) -> ConversionConfig:
        """Get conversion configuration"""
        return self.config.conversion

    def get_image_config(self) -> ImageConfig:
        """Get image configuration"""
        return self.config.image

    def get_mermaid_config(self) -> MermaidConfig:
        """Get Mermaid configuration"""
        return self.config.mermaid

    def get_llm_config(self) -> LLMConfig:
        """Get LLM configuration"""
        return self.config.llm

    def update_config(self, updates: Dict[str, Any]):
        """
        Update configuration with new values.

        Args:
            updates: Dictionary with config updates
        """
        if 'conversion' in updates:
            for key, value in updates['conversion'].items():
                if hasattr(self.config.conversion, key):
                    setattr(self.config.conversion, key, value)

        if 'image' in updates:
            for key, value in updates['image'].items():
                if hasattr(self.config.image, key):
                    setattr(self.config.image, key, value)

        if 'mermaid' in updates:
            for key, value in updates['mermaid'].items():
                if hasattr(self.config.mermaid, key):
                    setattr(self.config.mermaid, key, value)

        if 'llm' in updates:
            for key, value in updates['llm'].items():
                if hasattr(self.config.llm, key):
                    setattr(self.config.llm, key, value)

        logger.info("Configuration updated")


def create_default_config(output_path: Path, format: str = 'yaml'):
    """
    Create a default configuration file.

    Args:
        output_path: Path to save config file
        format: Output format ('yaml' or 'json')
    """
    config = MarkdownMakerConfig()
    manager = ConfigManager()
    manager.config = config
    manager.save_config(output_path, format)
    print(f"Created default config at: {output_path}")


def print_config(config: MarkdownMakerConfig):
    """Print configuration in a readable format"""
    print("Current Configuration:")
    print("=" * 60)

    print("\n[Conversion Settings]")
    print(f"  Object Threshold: {config.conversion.object_threshold}")
    print(f"  Output Directory: {config.conversion.output_dir}")
    print(f"  Enable Mermaid: {config.conversion.enable_mermaid}")
    print(f"  Save Metadata: {config.conversion.save_metadata}")
    print(f"  Create TOC: {config.conversion.create_toc}")

    print("\n[Image Settings]")
    print(f"  Dimensions: {config.image.width}x{config.image.height}")
    print(f"  DPI: {config.image.dpi}")
    print(f"  Quality: {config.image.quality}")
    print(f"  Thumbnail Size: {config.image.thumbnail_size}")

    print("\n[Mermaid Settings]")
    print(f"  Enabled: {config.mermaid.enabled}")
    print(f"  Enable Styling: {config.mermaid.enable_styling}")
    print(f"  Max Nodes: {config.mermaid.max_nodes}")
    print(f"  Auto Detect: {config.mermaid.auto_detect}")

    print("\n[LLM Settings]")
    print(f"  Enabled: {config.llm.enabled}")
    print(f"  Model: {config.llm.model}")
    print(f"  API Key: {'***' if config.llm.api_key else 'Not set'}")
    print(f"  Base URL: {config.llm.base_url or 'Default'}")
    print(f"  Temperature: {config.llm.temperature}")
    print(f"  Max Tokens: {config.llm.max_tokens}")

    print("\n" + "=" * 60)


if __name__ == "__main__":
    # Test configuration management
    print("Testing ConfigManager...\n")

    # Create default config
    manager = ConfigManager()
    print_config(manager.config)

    # Save to file
    output_path = Path("markdown_maker.example.yaml")
    manager.save_config(output_path, format='yaml')
    print(f"\nSaved example config to: {output_path}")
