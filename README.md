# dgfexporter
Exports "Fakturen" in "DIALOG Gebührenfakturierung" to PDF files

## Installation
1. Download and install [AutoIt](https://www.autoitscript.com/site/autoit/)
2. Download and install [FreePDF](http://freepdfxp.de/index_de.html) and [GhostScript](http://www.ghostscript.com/download/gsdnld.html)
2. Open the SciTE Script Editor
3. Open the file "GEB_autoPrint.au3"
4. Change the variables accordingly to your needs.
5. Open "Gebührenfakturierung" with a user which has GEB_Admin rights (preferably on the server itself)
6. Set the default printer settings
  1. Datei -> Drucker einrichten
  2. Ausgabe: Drucker
  3. Drucker: FreePDF
  4. Papierformat: A3 or A4
  5. Confirm with OK
7. Start the script in SciTE with Tools -> Go

## Variables
### restartEntry
Change this variable if you don't want to start with the first entry. If set to False the script will start at the beginning. e.g. "2015 Fakturen Woche 19"

### lastEntry
You have to change this variable. Set it to the last entry in the list "Serie". The name has to be exact!

### oPath
This is the export folder. All PDFs will be saved here. Make sure the user has full access rights.

### logFile
Change this variable to choose another name for the logfile.

### maxPDFperReport
A report can have multiple subreports. Only change this variable if you encounter problems with the script (e.g. unexpected stopping of the export)

### timeout
Change this if your machine is slow or your reports are very big. 30 seconds is normaly enough, shoulden't be set to more than 600 seconds.

### debug
WARNING: this variable can harm the machine!
Only change this variable if you're sure what you're doing! It can harm the machine, as the script will not check if the desired window is active! It will close all windows which don't belong to "Gebührenfakturierung"
