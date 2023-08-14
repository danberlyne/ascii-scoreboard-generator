#!/usr/bin/env python3
# ascii_scoreboard_generator.py - Generates a set of ASCII scoreboards from a spreadsheet of voting data.

import openpyxl
from pathlib import Path

def main():
    wb = openpyxl.load_workbook('voting_data.xlsx')
    ws = wb.active
    voters = [cell for cell in ws['1'] if cell.value != '' and cell.value != None]
    competitors = [cell for cell in ws['A'] if cell.value != '' and cell.value != None]

    max_name_length = max(max(len(voter.value) for voter in voters), max(len(competitor.value) for competitor in competitors))
    scoreboard_width = max_name_length + 14
    horizontal_line = ' ' + '_' * scoreboard_width + ' '
    dead_space = '|' + ' ' * scoreboard_width + '|'
    header = '|' + '%s' + '|'
    competitor_scorecard = '| ' + '%s' + '%s' + ' | ' + '%s' + ' |'
    footer = '|' + '_' * scoreboard_width + '|'
    # `template` has the following string interpolation:
    # - `voter`;
    # - `competitor`, `vote`, and `running_total` for each competitor.
    template = f'{horizontal_line}\n{dead_space}\n{header}\n{footer}\n{dead_space}\n' + f'{competitor_scorecard}\n' * len(competitors) + footer

    running_totals = [0 for competitor in competitors]
    for i, voter in enumerate(voters):
        votes = [vote.value for vote in ws[f'{voter.column_letter}'] if vote.value != voter.value]
        for j in range(len(votes)):
            if votes[j] == None:
                votes[j] = ''
        scoreboard_data = (f'NOW VOTING: {voter.value}'.center(scoreboard_width),)
        for j in range(len(running_totals)):
            if type(votes[j]) == int or type(votes[j]) == float:
                running_totals[j] += votes[j]
            scoreboard_data += (competitors[j].value.ljust(scoreboard_width - 10), str(votes[j]).rjust(2), str(running_totals[j]).zfill(3))
        scoreboard = template % scoreboard_data
        file = open(Path('Scoreboards', f'{i + 1}'.zfill(2) + f' {voter.value}.txt'), 'w')
        file.write(scoreboard)
        file.close()

if __name__ == '__main__':
    main()