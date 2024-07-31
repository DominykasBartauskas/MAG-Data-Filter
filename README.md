# MAG-Data-Filter
Some small code to enhance the points tallying from .xlsx documents.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Explanation](#explanation)
- [Contributing](#contributing)
- [Known Bugs](#known-bugs)
- [License](#license)

## Installation
1. Install Python from the [official website](https://www.python.org/downloads/). 
>[!NOTE]
>For this app, ```Python 3.12``` and above is necessary.
2. Clone the repository:
```
git clone https://github.com/DominykasBartauskas/MAG-Data-Filter.git
cd discord-data-filter
```
3. Set up a virtual environment:
```
python -m venv venv
venv\Scripts\activate
```
4. Install the required dependencies:
```
pip install -r requirements.txt
```

## Usage
1. Copy and paste the Excel (*.xlsx*) file (renamed to `data.xlsx`) with the data into the root directory of the project.
2. Run the script using the run button or this command `python app.py` in the terminal.
3. The script will generate an output Excel (*.xlsx*) file `processed_data.xlsx` with the following sheets:
- `Usernames Count`: Tally of usernames with their points. These can be directly copied and pasted into Discord. A `@` symbol before the username is required to create the mention in Discord.
- `Lost Rows`: Rows that didn't meet the criteria for counting. These have to be manually checked.

## Explanation
The script processes each row of the input Excel file and performs the following operations:
1. Checks if the row contains an image link in the link column and usernames in the `Mentions` column.
2. Extracts usernames from the `Mentions` column.
3. Analyzes the `Content` column to determine the result of the match and tally points accordingly.
4. Tallies points for usernames if the result indicates a win.
5. Moves rows that don't meet the criteria or have missing values to the Lost Rows sheet.

## Contributing
1. Fork the repository.
2. Create a new branch using `git checkout -b feature-branch`.
3. Make your changes and commit them using `git commit -am 'Add new feature'`.
4. Push to the branch using `git push origin feature-branch`.
5. Create a new Pull Request.

## Known Bugs
- The Discord link breaks after moving it into the new Excel (*.xlsx*) file.

## License
This project is licensed under the MIT License.