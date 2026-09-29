from datetime import datetime
from pathlib import Path

from config_loader import load_config


def is_valid_date_folder(folder_path, date_format):
    """Check whether a folder name follows the configured date format."""
    try:
        datetime.strptime(folder_path.name, date_format)
        return True
    except ValueError:
        return False


def discover_input_files(config):
    """Find valid Excel files inside date-wise input folders."""

    input_folder = Path(config["project"]["input_folder"])

    accepted_extensions = {
        extension.lower()
        for extension in config["files"]["accepted_extensions"]
    }

    date_format = config["files"]["date_folder_format"]

    if not input_folder.exists():
        raise FileNotFoundError(
            f"Input folder does not exist: {input_folder}"
        )

    if not input_folder.is_dir():
        raise NotADirectoryError(
            f"Configured input path is not a directory: {input_folder}"
        )

    date_folders = [
        folder
        for folder in input_folder.iterdir()
        if folder.is_dir()
        and is_valid_date_folder(folder, date_format)
    ]

    date_folders.sort(
        key=lambda folder: datetime.strptime(
            folder.name,
            date_format,
        )
    )

    discovered_files = []

    for date_folder in date_folders:
        for file_path in sorted(
            date_folder.rglob("*"),
            key=lambda path: str(path).casefold(),
        ):
            if not file_path.is_file():
                continue

            # Ignore temporary Excel lock files.
            if file_path.name.startswith("~$"):
                continue

            if file_path.suffix.lower() not in accepted_extensions:
                continue

            discovered_files.append(
                {
                    "date_folder": date_folder.name,
                    "file_name": file_path.name,
                    "file_path": str(file_path.resolve()),
                    "file_size_bytes": file_path.stat().st_size,
                }
            )

    return discovered_files


if __name__ == "__main__":
    settings = load_config()
    files = discover_input_files(settings)

    print("Input file discovery completed")
    print(
        f"Input folder: "
        f"{settings['project']['input_folder']}"
    )
    print(
        f"Valid Excel files found: {len(files)}"
    )

    for item in files:
        print(
            f"{item['date_folder']} | "
            f"{item['file_name']} | "
            f"{item['file_path']}"
        )
