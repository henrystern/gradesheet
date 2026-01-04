"""Create a marking sheet from a class list and rubric."""

import argparse
from pathlib import Path
from typing import List

from loguru import logger
import polars as pl
import xlsxwriter

from gradesheet.config import (
    CLASSLIST_DIR,
    LOGS_DIR,
    PROCESSED_DATA_DIR,
    PROJ_ROOT,
    RUBRICS_DIR,
)


def main(args: None | List = None):
    """Create a marking sheet from a class list and rubric."""
    parser = argparse.ArgumentParser(
        prog="Create grade sheet",
        description=__doc__,
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--class-list",
        type=Path,
        help="Path to class list CSV file.",
    )
    parser.add_argument(
        "--rubric",
        type=Path,
        help="Path to rubric CSV file.",
    )
    parsed_args = parser.parse_args(args)

    if parsed_args.class_list is None:
        parsed_args.class_list = prompt_file_selection(CLASSLIST_DIR, ".csv")
    if parsed_args.rubric is None:
        parsed_args.rubric = prompt_file_selection(
            RUBRICS_DIR, ".csv", recursive=True
        )

    make_marking_sheet(
        output_path=PROCESSED_DATA_DIR
        / f"{parsed_args.class_list.stem} - {parsed_args.rubric.stem} - marking.xlsx",
        rubric=pl.read_csv(parsed_args.rubric),
        class_list=pl.read_csv(parsed_args.class_list),
    )


def make_marking_sheet(
    output_path: Path,
    rubric: pl.DataFrame,
    class_list: pl.DataFrame,
    id_cols: List = ["Last Name", "First Name", "OrgDefinedId"],
):
    """Make a long format sheet for marking and comments."""
    df = (
        class_list.with_columns(
            pl.lit(None).alias(c) for c in (rubric.get_column("Question"))
        )
        .unpivot(
            index=id_cols,
            value_name="Deductions",
            variable_name="Question",
        )
        .with_columns(pl.lit(None).alias("Comment"))
        .sort(id_cols + ["Question"])
    )

    formulas = {}
    formulas["Max score"] = {
        "formula": '=XLOOKUP([@Question],Rubric[Question],Rubric[Max score], "", 0)',
        "insert_after": "Question",
    }
    formulas["Score"] = {
        "formula": '=IF(ISBLANK([@Deductions]), "",[@[Max score]]-[@Deductions])',
        "insert_after": "Comment",
    }
    if "Folder" in id_cols:
        formulas["Hyperlink"] = {
            "formula": '=HYPERLINK("./" & [@Folder], "Link")',
            "insert_after": "Folder",
        }

    with xlsxwriter.Workbook(output_path) as workbook:
        rubric.write_excel(
            workbook,
            worksheet="Rubric",
            table_name="Rubric",
            autofit=True,
        )
        df.write_excel(
            workbook,
            autofit=True,
            freeze_panes=(1, 1),
            formulas=formulas,
            worksheet="Marking",
            table_name="Marks",
        )
    logger.success(
        f"Wrote marking sheet to {output_path.relative_to(PROJ_ROOT)}"
    )
    print("Score validation formula:")
    print('=INDIRECT("RC[-1]", FALSE)')
    print("Progress formula:")
    print(
        "=LET(total, AGGREGATE(3, 5, [Question]), filled, AGGREGATE(3, 5, [Deductions]), filled / total)"
    )
    print("pivot table Average formula:")
    print("=Score /'Max score'")


def prompt_file_selection(
    directory: Path,
    file_extension: str = ".csv",
    recursive: bool = False,
):
    """Prompt the user to select a file from a directory."""
    print(f"Please select a file from {directory.relative_to(PROJ_ROOT)}:")
    path_prefix = "**/" if recursive else ""
    choices = sorted(directory.glob(f"{path_prefix}*{file_extension}"))
    for idx, choice in enumerate(choices, 1):
        print(f"{idx}: {choice.relative_to(directory)}")
    selection = int(input("Enter the number of your choice: ")) - 1
    return choices[selection].resolve()


if __name__ == "__main__":
    logger.add(LOGS_DIR / "create.log")
    main()
