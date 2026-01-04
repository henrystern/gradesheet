"""Clean the brightspace submissions by (for example) removing duplicate submissions, converting to pdf, and renaming folders."""

from pathlib import Path
import shutil
from typing import Iterator

import docx2pdf
import polars as pl

from gradesheet.config import PROCESSED_DATA_DIR, RAW_DATA_DIR


def convert_docs_to_pdf():
    """Convert the .docx submissions to pdfs keep other documents unchanged."""
    for f in RAW_DATA_DIR.rglob("*"):
        if f.is_dir():
            continue
        relative_path = f.relative_to(RAW_DATA_DIR)
        output_file = PROCESSED_DATA_DIR / relative_path
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if f.suffix == ".docx" and not f.stem.startswith("-$"):
            docx2pdf.convert(f, output_file.with_suffix(".pdf"))
        else:
            shutil.copyfile(f, output_file)


def keep_latest_submission():
    """Only keep the latest submission from each student and move it from RAW_DATA_DIR to PROCESSED_DATA_DIR."""
    submissions = [str(d) for d in RAW_DATA_DIR.glob("*") if d.is_dir()]
    df = (
        pl.DataFrame(submissions, ["Folder"])
        .with_columns(
            pl.col("Folder")
            .str.split_exact(" - ", 3)
            .struct.rename_fields(["ID", "Name", "Submitted"])
            .alias("Split")
        )
        .unnest("Split")
        .drop("ID")
        .with_columns(
            pl.col("Submitted")
            .str.replace(r" (\d\d\d [A-Z]*)$", " 0${1}")  # Zero pad hours
            .str.to_datetime("%b %d, %Y %I%M %p")
        )
        .sort(["Submitted", "Name"])
        .unique("Name", keep="last", maintain_order=True)
    )
    for folder in df["Folder"]:
        relative_path = Path(folder).relative_to(RAW_DATA_DIR)
        output_folder = PROCESSED_DATA_DIR / relative_path
        output_folder.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(folder, output_folder)


def get_submission_metadata(submissions_dir: Iterator[Path]):
    """Get the metadata for each submission from the submission folder names."""
    submissions = [d.stem for d in submissions_dir if d.is_dir()]
    return (
        pl.DataFrame(submissions, ["Folder"])
        .with_columns(
            pl.col("Folder")
            .str.split_exact(" - ", 3)
            .struct.rename_fields(["ID", "Name", "Submitted"])
            .alias("Split")
        )
        .unnest("Split")
        .drop("ID")
        .with_columns(
            pl.col("Submitted")
            .str.replace(r" (\d\d\d [A-Z]*)$", " 0${1}")  # Zero pad hours
            .str.to_datetime("%b %d, %Y %I%M %p")
        )
        .sort(["Submitted", "Name"])
        .unique("Name", keep="last", maintain_order=True)
    )


if __name__ == "__main__":
    keep_latest_submission()
    submissions = get_submission_metadata(PROCESSED_DATA_DIR.glob("*"))
