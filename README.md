# Gradesheet

Create a spreadsheet for grading assignments and tests.

## Project structure

```
│   LICENSE                               # MIT license
│   pyproject.toml                        # Package metadata and configuration
│   README.md                             # Brief documentation
├───data
│   ├───processed                         # The output spreadsheets. 
│   └───raw                               # The input data.
└───src
    └───gradesheet                        # Python package for this project
            __init__.py                   # Makes package installable
            config.py                     # Configuration parameters
            create.py                     # Script to create the gradesheet
```


## Python environment

The python environment is managed by [uv](https://docs.astral.sh/uv/getting-started/installation/).

Once uv is installed, create the environment by running:

```bash
uv sync
```

Then activate the environment with:

```bash
source .venv/bin/activate
```

If you are using vscode, set your python interpretor to `.venv/bin/python`.

The clean_submissions module requires optional dependencies that can be installed via the submissions group:

```bash
uv sync --extra submissions
```

## Usage

To create a gradesheet, run the following command:

```bash
uv run src/gradesheet/create.py
```

You can optionally pass the class list and rubric file paths as arguments. If not provided the script will prompt you to choose a file from the files in `data/raw/`. 

The input files must match the format shown in the examples in `data/raw/classlists/classlist_demo.csv` and `data/raw/rubrics/rubric_demo.csv`.

The script will output the gradesheet to `data/processed` with a filename based on the name of the class list and rubric files used. The output filename will be in the format `<class_list_name> - <rubric_name> - marking.xlsx`.

Some features of the gradesheet (data validation and progress tracking) need to be configured from within excel. For that purpose a macro is provided in `src/vba/SetupGradesheet.vba`. To use the macro import it into your PERSONAL.XLSB workbook and run it immediately after opening the generated gradesheet.

An example output is provided in `data/processed/classlist_demo - rubric_demo - marking.xlsx`. This shows the gradesheet after running the setup macro.

You fill out the gradesheet by entering the deducted marks and any comments in each row. The final grades are calculated in the `Grades` pivot table.