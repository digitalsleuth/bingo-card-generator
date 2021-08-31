# bingo-card-generator
Interactive Bingo Card and PDF Generator

Tool to generate a clickable bingo card for virtual bingo events, and generate PDF's for printing these cards
Requires pdfkit and wkhtmltopdf for OS it's being run on (Windows requires an exe, linux install from pkg manager)

This script currently will only generate 6 cards on one sheet (2 rows of 3 cards).

While not normally necessary for CSS, the float values are required for the proper printing of the PDF's
If these are removed/modified, you can expect your PDF's to be off-center or misaligned.

Icon made by Freepik (https://www.freepik.com) from Flaticon (www.flaticon.com)

Standalone executable made using pyinstaller:
`pyinstaller -F --icon=bingo.ico bingo-card-generator.py`

TO DO:
- Make number of cards per sheet customizable
