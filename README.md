
# Excel MCQ Validator

This script processes an Excel sheet containing multiple-choice questions (MCQs). Each row in the sheet represents a question with four options and the correct answer. The script evaluates the provided answer and marks it as 'Correct' or 'Wrong' in the last cell of each row.

## Requirements

- Python 3.x
- `openpyxl` library

## Installation

1. Install Python 3.x from the official [Python website](https://www.python.org/).
2. Install the `openpyxl` library using pip:

```sh
pip install openpyxl

Usage
Prepare an Excel file with the following structure:

Each row should contain:
Question
Option 1
Option 2
Option 3
Option 4
Correct Answer (one of the options)
The first row should contain headers and will be skipped by the script.
Run the script:

The script will process each row, evaluate the answer, and append 'Correct' or 'Wrong' in the last cell of each row.
Example
Given an Excel file mcq_questions.xlsx with the following content:

Question	Option 1	Option 2	Option 3	Option 4	Correct Answer
What is 2+2?	3	4	5	6	4
Capital of USA	New York	Washington D.C.	Los Angeles	Chicago	Washington D.C.
After running the script, the file will be updated to:

Question	Option 1	Option 2	Option 3	Option 4	Correct Answer	Result
What is 2+2?	3	4	5	6	4	Correct
Capital of USA	New York	Washington D.C.	Los Angeles	Chicago	Washington D.C.	Correct
License
This project is licensed under the MIT License.

Author
[Nitin Gopi]


