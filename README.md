# PDSBot script
### AUTOMATED BULK. PDS TO CONVERTER
### Demo version for GitHub
### Developed by: Jakub Pitera

PDSBot is a standalone python-based tool that takes control over your machine allowing for automated conversion of .PDS files to .TIF image format using Schlumberger PDSView software.

In order to convert make sure:
 - "Schlumberger PDSView" is installed and running.
 - in 'File' > 'Options', "Display message box after PDS file has been loaded" is ticked.
 
Specific for 'Print to PDF' conversion:
 - in 'File' > 'Print Setup', "Microsoft Print to PDF" is selected
 - in 'File' > 'Print', "Show print status" is unticked

In PDSView select desired DPI of .TIF and any other PDF printing options manually before converting!

There are two ways of converting to PDF: 'File > Save as...' and 'File > Print'. The latter is using 'Microsoft Print to PDF'.
Use it if the 'File' > 'Save  as...' > 'Save as type:' > 'PDF Files (*.PDF)' is unavailable in PDSView.

To start bulk conversion, please first select MODE using the prompt below. 
Secondly, enter the directory of .PDS files.
All the .PDS files inside that directory tree will be converted (including the subfolders).

Important note: This program will take control of the mouse and keyboard until the conversion is finished. 
To interrupt, please hold "CTRL"+"C" while this window is active. 
You can return to the interrupted conversion by selecting "4 CONTINUE from saved log" MODE and then providing 
a full path of a saved log created by POT Bot.
