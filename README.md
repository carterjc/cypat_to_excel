# cypat_to_excel
Tool that scrapes the CyberPatriot public scoreboard and dumps the data into an excel file.

# Synopsis
This program will generate a detailed Excel file that can be used as one pleases. Be warned, this program takes a particularly long time to run (especially in the early rounds) because it makes thousands of individual GET requests to discover individual image information.

# How to Run
First, clone the repository:
`git clone https://github.com/carterjc/cypat_to_excel.git`

Navigate to the folder:
`cd cypat_to_excel`

Ensure the following modules are installed:
- xlsxwriter
- bs4
- requests

If needed, use: `pip install ______` (module)

Run the script with `python main.py` and enter accurate data into the requested fields.

After a bit of waiting, enjoy!
