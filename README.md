# Bingo Card Generator
## bingo-card-generator
Interactive and PDF Bingo Card Generator

Tool to generate a clickable, interactive bingo cards for virtual bingo events, and generate PDF's for offline use (for those who would rather print and use paper).  

Requires pdfkit, Pillow, and wkhtmltopdf (wkhtmltopdf is dependent on OS it's being run on (Windows requires an exe, linux install from pkg manager))

This tool generates 6 cards on one sheet (2 rows of 3 cards).

Application (exe) Icon made by Freepik (https://www.freepik.com) from Flaticon (www.flaticon.com)

## Options  

There are many options available:

Choose a custom logo to use as a dauber:  
- when chosen, a JPG or PNG must be supplied, and it will be convered to 48x48  
- for best results, it should be square to begin with  

Allow the player to choose their own dauber during play:  
- when allowed, the player can choose their own digital dauber for each game  
- cannot be changed after selected during gameplay, to avoid confusion  
- once card is "cleared", a new dauber can be chosen  
- option to allow, not allow can be defined prior to card generation  
  
Allows the generator of the card to choose the card colour and, if applicable, the dauber colour:  
- depending on the dauber, the colour can be chosen, but some daubers cannot be changed  
- the circle, square, maple leaf, and heart can be colour-adjusted  
  
Generates a spreadsheet to track all numbers called, and which cards actually have those numbers:  
- gives the ability to confirm an actual "BINGO" quickly, by selecting the card number and visually identifying the "shape"
- when numbers are called, type them into the "CALL" sheet, and they will highlight on each card number  
  
## Examples
### Type in the colour name

![Type-In-Colour](https://user-images.githubusercontent.com/62841822/216872120-68d360b5-57b2-4d82-82a1-dd954a1387ee.png)  
![Type-In-Colour-2](https://user-images.githubusercontent.com/62841822/216872182-7d0e6047-de17-4b3c-9453-f6553998e895.png)  

### Use the Colour Picker  

![Color-Picker](https://user-images.githubusercontent.com/62841822/216872293-20a63e1b-87c9-4ef7-9913-dde91da498c1.png)


### Choose your own logo  
![ChooseLogo](https://user-images.githubusercontent.com/62841822/216872424-0b61be79-d4d1-4530-bf81-3cc4cc8d148b.png)  
![ChooseLogo-2](https://user-images.githubusercontent.com/62841822/216872447-313497f9-541d-4552-baf2-0b57b69f768f.png)

### Allow the players to choose their own dauber  
![ChooseDauber](https://user-images.githubusercontent.com/62841822/216872552-18e31d86-42ff-488e-ab81-5bf096aee78a.png)


### Standalone Executable
Standalone executable made using pyinstaller:
`pyinstaller -F --icon=bingo.ico bingo-card-generator.py`

