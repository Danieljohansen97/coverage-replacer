# Coverage Replacer 2023
- SVV.xlsx shows roadnumber changes sorted by Fylkesnummer from 1946-2020, so the old fylkesnummer. 
- Currently our Coverage Descriptions are sorted by kommune.
- The program then maps the kommunes to the old list of Fylker, and then looks through Coverage Descriptions looking for the said kommune and updates the road number based on which fylkesnummer it has, as the SVV.xlsx uses those fylkesnummer for ordering.

## How to run

`code` $ python coverage-replacer.py `code`

---
Made by Daniel Johansen 2023# coverage-replacer
