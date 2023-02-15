# Bingo Card Generator

## Generate interactive BINGO cards and PDF's for printing!

Tool to generate a clickable, interactive bingo cards for virtual bingo events, and generate PDF's for offline use (for those who would rather print and use paper).  

Requires pdfkit, Pillow, webcolors, and wkhtmltopdf (wkhtmltopdf is dependent on OS it's being run on (Windows requires an exe, linux install from pkg manager))  

Make sure you get wkhtmltopdf first before running Bingo Card Generator - [Linux and Windows Flavours](https://github.com/wkhtmltopdf/packaging/releases/tag/0.12.6-1)

This tool generates 6 cards on one sheet (2 rows of 3 cards).

Application (exe) Icon made by Freepik (https://www.freepik.com) from Flaticon (www.flaticon.com)

## Options  

There are many options available:

### Choose a custom logo to use as a dauber  
- when chosen, a JPG or PNG must be supplied, and it will be convered to 48x48  
- for best results, it should be square to begin with  

### Allow the player to choose their own dauber during play  
- when allowed, the player can choose their own digital dauber for each game  
- cannot be changed after selected during gameplay, to avoid confusion  
- once card is "cleared", a new dauber can be chosen  
- option to allow, not allow can be defined prior to card generation  
  
### Allows the generator of the card to choose the card colour and, if applicable, the dauber colour  
- depending on the dauber, the colour can be chosen, but some daubers cannot be changed  
- the circle, square, maple leaf, and heart can be colour-adjusted  
  
### Generates a spreadsheet to track all numbers called, and which cards actually have those numbers  
- gives the ability to confirm an actual "BINGO" quickly, by selecting the card number and visually identifying the "shape"
- when numbers are called, type them into the "CALL" sheet, and they will highlight on each card number  
  
### Easy mode: When selected from the Mode menu, will make 'daubbing' your numbers easier
- In Easy Mode, if B4 is called, clicking on it once on one card will select it for all cards  
- When not selected, if B4 is called, you will have to select it on each card manually  

### Toggle the dauber
- Click the wrong number? Click it again to remove the dauber!

## Examples  
### Type in the colour name


![Select-Bingo-Card-Options](https://user-images.githubusercontent.com/62841822/218927016-3f07aaaf-8fe3-4aa5-b673-52e8461341ad.png)

![Lime-Maple-Leaf-Example](https://user-images.githubusercontent.com/62841822/218927613-a9ea6036-d591-4df9-bc10-58b54f37db93.png)

### Use the Colour Picker  

![Use-Colour-Picker](https://user-images.githubusercontent.com/62841822/218927839-c4189bc8-61ce-4039-904d-ead4da039c6f.png)

### Choose your own logo  
![Choose-Logo](https://user-images.githubusercontent.com/62841822/218928230-40a6c420-03f7-4f26-af68-e82f1c597617.png)  
![ChooseLogo-2](https://user-images.githubusercontent.com/62841822/216872447-313497f9-541d-4552-baf2-0b57b69f768f.png)

### Allow the players to choose their own dauber  
![ChooseDauber](https://user-images.githubusercontent.com/62841822/216872552-18e31d86-42ff-488e-ab81-5bf096aee78a.png)


