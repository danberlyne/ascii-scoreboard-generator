# About
This tool generates scoreboards using ASCII art, based on voting data entered in the spreadsheet `voting_data.xlsx`. One scoreboard is generated for each voter, containing the voter's name, their votes, and running totals of the competitors' scores.

# Requirements
This tool uses the third party module OpenPyXL. Please install this module before attempting to run the Python scripts included here.

# Instructions
1. Enter the competitor names, voter names, and all sets of votes in the spreadsheet `voting_data.xlsx`.
2. Run `ascii_scoreboard_generator.py`. This will generate scoreboards using the data in `voting_data.xlsx`, saving them to the `Scoreboards` folder as `.txt` files.