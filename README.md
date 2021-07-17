# bingo-card-generator
Interactive Bingo Card Generator

Tool to generate a clickable bingo card for virtual bingo events, and generate PDF's for printing these cards
Requires pdfkit and wkhtmltopdf for OS it's being run on (Windows requires an exe, linux install from pkg manager)

This script currently will only generate 6 cards on one sheet (2 rows of 3 cards).
Hex codes can be used in place of words for colours, but must be wrapped in quotes on the command line

While not normally necessary for CSS, the float values are required for the proper printing of the PDF's
If these are removed/modified, you can expect your PDF's to be off-center or misaligned.
