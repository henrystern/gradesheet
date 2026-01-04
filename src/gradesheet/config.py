"""Configuration parameters for convertdoc."""

from pathlib import Path

from loguru import logger

# Paths
PROJ_ROOT = Path(__file__).resolve().parents[2]
"""The project root directory. All other directories should be defined relative to this root."""

DATA_DIR = PROJ_ROOT / "data"
RAW_DATA_DIR = DATA_DIR / "raw"
CLASSLIST_DIR = RAW_DATA_DIR / "class_lists"
RUBRICS_DIR = RAW_DATA_DIR / "rubrics"
INTERIM_DATA_DIR = DATA_DIR / "interim"
PROCESSED_DATA_DIR = DATA_DIR / "processed"

LOGS_DIR = PROJ_ROOT / "logs"

# If tqdm is installed, configure loguru with tqdm.write
try:
    from tqdm import tqdm

    logger.remove(0)
    logger.add(lambda msg: tqdm.write(msg, end=""), colorize=True)
except (ModuleNotFoundError, ValueError):
    pass
