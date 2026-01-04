"""Evaluate student grades and generate a report."""

from itertools import product
from pathlib import Path

import polars as pl
import xlsxwriter

from gradesheet.config import PROCESSED_DATA_DIR


def pivot_with_comments(
    path: Path, long: pl.DataFrame, num_index_columns: int = 5
):
    """Pivot the long format grading sheet to a wide format with overlaid comments."""
    scores = long.drop("Comment").pivot(on="Question", values="Score")
    comments = long.drop("Score").pivot(on="Question", values="Comment")
    with xlsxwriter.Workbook(path) as workbook:
        worksheet = workbook.add_worksheet("Grades")
        scores.write_excel(workbook=workbook, worksheet="Grades")
        for i, j in product(
            range(comments.height),
            range(num_index_columns, comments.width),
        ):
            if comments[i, j] is None:
                continue
            worksheet.write_comment(i + 1, j, comments[i, j])


if __name__ == "__main__":
    pivot_with_comments(
        path=PROCESSED_DATA_DIR / "grades.xlsx",
        long=pl.read_excel(
            PROCESSED_DATA_DIR / "marking.xlsx", sheet_name="Marking"
        ).drop("Max score"),
    )
