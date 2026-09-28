from pathlib import Path

import yaml


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config" / "settings.yaml"


def load_config(config_path=DEFAULT_CONFIG_PATH):
    """Load and validate the pipeline configuration file."""

    config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_path}"
        )

    with config_path.open("r", encoding="utf-8") as file:
        config = yaml.safe_load(file)

    if not isinstance(config, dict):
        raise ValueError("The configuration file is empty or invalid.")

    required_sections = [
        "project",
        "files",
        "columns",
        "units",
        "cleaning",
        "pairing",
        "cropping",
        "phases",
        "traffic",
        "outputs",
        "processing",
    ]

    missing_sections = [
        section
        for section in required_sections
        if section not in config
    ]

    if missing_sections:
        raise ValueError(
            "Missing configuration sections: "
            + ", ".join(missing_sections)
        )

    return config


if __name__ == "__main__":
    settings = load_config()

    print("Configuration loaded successfully")
    print(f"Input folder: {settings['project']['input_folder']}")
    print(f"Results folder: {settings['project']['results_folder']}")
    print(
        "Inching speed range: "
        f"{settings['phases']['inching_min_speed_kmh']} to "
        f"{settings['phases']['inching_max_speed_kmh']} km/h"
    )