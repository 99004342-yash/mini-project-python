# Automated excel with Python 3

- Can be used for extracting Ps number data from an excel sheet.
- Highly dynamic except:
1. Your excel file must contain PS Numbers of candidate in first column.
2. Custom excel files should be placed inside `src/` folder with name of `input.xlsx`.

## Instructions

#### For custom excel input

- **For your custom input excel, follow the following steps:** 
1. Place your excel file within `src/` folder.
2. CHange the name of your excel to `input.xlsx`.
3. Run the program following the instructions stated below.

##### Note
- YOur excel file must contain PS Numbers of candidate in first column.

### Steps to use:

**Make sure you are inside _`/src`_ folder before running the command**

1. run `pip3 install -r requirements.txt`
2. run `cd ./src`
3. run `python3 index.py`

### Pylint score - 9.52

**Make sure you are inside _`/src`_ folder before running the command**

- run `pylint index.py`

### Pytest

**Make sure you are inside _`/src`_ folder before running the command**

- run `pytest tests.py`

[![Tests](https://github.com/99004342-yash/mini-project-python/actions/workflows/python-package.yml/badge.svg)](https://github.com/99004342-yash/mini-project-python/actions/workflows/python-package.yml)


