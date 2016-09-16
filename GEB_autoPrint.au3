#include <Constants.au3>
#include <Date.au3>
#include <Array.au3>
#include <File.au3>

;
; AutoIt Version: 3.0
; Language:       English
; Platform:       WinNT
; Author:         Stefan Blatter (sbl at brienz dot ch)
;
; Script Function:
;   Exports "Fakturen" in "DIALOG Gebührenfakturierung" to PDF files
;   Tested with version 5.33.008
;
;   - included print modules:
;     - FreePDF 4.14

;;; Log IDs
; 0000 = Default
; 0001 = Statistics
; 0404 = Template not found
; 0301 = Nothing to print (empty series)
; 0302 = Filepath too long, shortening
; 0101 = Start
; 0102 = End
; 0201 = Exported to Printer
; 0202 = File saved
; 0203 = Export to PDF finished
; 0204 = Multiple reports
; 0205 = Start processing serie
; 0501 = Unknown error (probably fatal)


; Edit these variables
Local $restartEntry = False					                    	; Last exported entry, continues with the next one (set False to start at the beginning)
Local $lastEntry = "2XXX Fakturen Woche 19" 						; Last series (exact name!)

Local $oPath = "X:\export"		 				                    ; Output path (the user needs full access to this folder. No "\" at the end!)

Local $logFile = $oPath & "\log.csv" 			                    ; Filename for logfile
Local $maxPDFperReport = 5 						                    ; The number of reports that can be in a single serie
Local $timeout = 30 						                    	; Max allowed timeout per serie (in seconds)

; Edit this variable at your own risk!
Global $debug = False

; Don't edit these variables
Local $totalEntrys = 0
Local $printedEntrys = 0

init($totalEntrys, $printedEntrys)

; Initialisation
Func init(ByRef $totalEntrys, ByRef $printedEntrys, $restartEntry = $restartEntry, $lastEntry = $lastEntry)
   myLog("Starte Export...", "101")

   WinActivate("[CLASS:ThunderRT6MDIForm]","")

   If WinActive("Gebührenfakturierung ","&Übermitteln") Then
	  Send("!c")
   EndIf

   ; Waits for window "Gebührefakturierung" to be active, then opens "Fakturen übermitteln..."
   WinWaitActive("[CLASS:ThunderRT6MDIForm]","")
   Send("!rf")
   WinWaitActive("Gebührenfakturierung ","&Übermitteln")

   ; Select "verbuchte"
   Send("{TAB}")
   Send("{RIGHT}")

   WinWaitActive("Gebührenfakturierung ","&Übermitteln")

   ; Select "Serie"
   Send("{TAB}{TAB}{TAB}{TAB}{TAB}")

   ; Restart after the last known position (if set)
   Local $sText = ""
   If $restartEntry == "" Or $restartEntry == False Then
	  Send("{DOWN}")
	  $totalEntrys = $totalEntrys + 1
   Else
	  myLog("Starte bei letztem erfolgreichen Eintrag neu")
	  Do ; Iterate through the list and searches for the last successfull entry
		 Send("{DOWN}")
		 $sText = ControlGetText("Gebührenfakturierung ","&Übermitteln", "ThunderRT5ComboBox1")
		 $totalEntrys = $totalEntrys + 1
	  Until $sText == $restartEntry
   EndIf

   ; Call handleReport to process the series
   Local $numPDFs = handleReport($totalEntrys)

   ; Log some statistics
   myLog("Total Serien: " & $totalEntrys, "1")
   myLog("Gespeicherte PDFs: " & $numPDFs, "1")
   myLog("Export abgeschlossen", "102")

   MsgBox(64,"Export abgeschlossen","Der Export wurde abgeschlossen."&@CRLF&@CRLF&"Total Serien: "&$totalEntrys &@CRLF&"Gespeicherte PDFs: "&$numPDFs )
EndFunc

; Select and prepare the report for export
Func handleReport(ByRef $totalEntrys, $lastEntry = $lastEntry, $maxPDFperReport = $maxPDFperReport, $oPath = $oPath, $timeout = $timeout)
   Local $numPDFs = 0
   Local $numPDFsTemp = 0

   Local $sleepTime = 50

   DO
	  ; Select the next entry
	  Send("{DOWN}")

	  ; Read text value of field "Serie"
	  Local $sText = ControlGetText("Gebührenfakturierung ","&Übermitteln", "ThunderRT5ComboBox1")

	  myLog($sText & " wird verarbeitet", "205")

	  ; Press "Drucken" button
	  Send("!d")

	  ; Checks for the required window
	  Local $c = 0
	  do
		 sleep($sleepTime)
		 If WinExists("Gebührenfakturierung ", "Druck aufbereiten") Then

		 Else
			$c = $c + 1
		 EndIf
		 ; If no required window exists after 2 minutes, continue and try to fix the problem automaticly. Log problem initiator
	  until WinExists("Serie drucken","Die Serie beinhaltet keine Rechnungen!") or WinExists("Drucken","Druckrichtung") or WinExists("Drucken von Fakturen","OK") or $c >= ($timeout*1000/$sleepTime)

	  ; Handle the active window
	  If WinExists("Serie drucken","Die Serie beinhaltet keine Rechnungen!") Then
		 ; No entrys found in this series, abort and continue with next entry
		 myLog($sText & " beinhaltet keine Rechnungen!", "301")
		 Send("{ENTER}")
		 WinWaitActive("Gebührenfakturierung ","&Übermitteln")
		 Send("+{TAB}+{TAB}+{TAB}") ; Select "Serie"
	  ElseIf WinExists("Drucken von Fakturen","OK") Then
		 ; Template is no longer available, report can't be generated, abort and continue with next entry
		 myLog("Template für " & $sText & " konnte nicht gefunden werden!", "404")
		 Send("{ENTER}")
		 WinWaitActive("Gebührenfakturierung ","&Übermitteln")
		 Send("+{TAB}+{TAB}+{TAB}") ; Select "Serie"
	  ElseIf WinExists("Drucken","Druckrichtung") Then
		 ; Everything ok, start to export this series
		 myLog($sText & " wird an Drucker übergeben", "201")
		 WinWaitActive("Drucken","Druckrichtung")
		 $numPDFsTemp = printReport()
		 $numPDFs = $numPDFs + $numPDFsTemp

		 Sleep(250)

		 ; Export report for each subreport
		 Local $counter = 0
		 do
			If WinExists("Drucken","Druckrichtung") Then
			   myLog($sText & " enthält mehrere Reports! Durchgang: "&$counter, "204")
			   $numPDFsTemp = printReport()
			   $numPDFs = $numPDFs + $numPDFsTemp
			EndIf
			$counter = $counter + 1
			sleep(50)
		 until WinExists("Drucken","Druckrichtung") or $counter >= $maxPDFperReport

		 ; Wait until required window is active again
		 WinWaitActive("Gebührenfakturierung ","&Übermitteln")

		 myLog($sText & " wurde abgeschlossen", "203")

		 Send("+{TAB}+{TAB}+{TAB}") ; select "Serie"
	  Else
		 If $debug == True Then
			; If everything fails, log it, try to fix it and continue with next entry
			Local $window = False
			Local $check = False
			Do
			   $window = WinGetHandle("")
			   If _WinGetClass($window) == "ThunderRT6MDIForm" Then
				  $check = True
			   Else
				  WinClose($window)
				  WinWaitClose($window)
				  $check = False
			   EndIf

			Until $check == True

			myLog("Something went terribly wrong. Tried to fix it! Initiator was " & $sText & ". Skipped!", "501","error")
		 Else
			myLog("Something went terribly wrong. DID NOT try to fix it! Initiator was " & $sText & ". Script aborted!", "501","error")
			MsgBox(16,"ERROR","Programm hat zu lange kein erwartetes Fenster geöffnet. Abbruch wegen Zeitüberschreitung."&@CRLF&"Der Fehlerhafte Eintrag war: "&$sText)
			Exit
		 EndIf

	  EndIf

	  $totalEntrys = $totalEntrys + 1
   Until $sText == $lastEntry

   ; Return the total amount of exported PDF files
   return $numPDFs
EndFunc


; Print the report
Func printReport($oPath = $oPath)
   Local $sText = ControlGetText("Gebührenfakturierung ","&Übermitteln", "ThunderRT5ComboBox1")

   Send("{ENTER}") ; Press the "OK" button

   WinWaitClose("Drucken","Druckrichtung")

   ; Call the specific PDF output method function (this one is for FreePDF XP)
   Local $numPDFs = printerFreePDF($sText, 3)

   return $numPDFs
EndFunc

Func printerFreePDF($filename, $max = 5, $oPath = $oPath)
   Local $k = 0
   Local $numPDFs = 0
   Do
	  WinWaitClose("Drucken von Datensätzen","Drucken abbrechen")

	  ; Wait until FreePDF is ready
	  WinWaitActive("FreePDF 4.14","&Ablegen",1)
	  If WinExists("FreePDF 4.14","&Ablegen") Then
		 WinActivate("FreePDF 4.14","&Ablegen")

		 ; Generate filepath
		 $filename = removeIllegalChars($filename)
		 Local $path = $oPath & "\" & $filename & "_" & getTime()
		 If StringLen($path) >= 256 Then ; Shorten filepath if it's too long
			$path = StringLeft($path, 256)
			myLog("Dateiname zu lang, er wurde gekürzt","302")
		 EndIf
		 $path = $path & ".pdf"

		 ; Uncheck "PDF anzeigen" and press the "Ablegen" button
		 ControlFocus("FreePDF 4.14","&Ablegen", 6)
		 Send("{SPACE}")
		 Send("!a")

		 ; Wait for the save dialog, insert full filepath and press the "Speichern" button
		 WinWaitActive("Speichern unter","&Speichern")
		 ControlSetText("Speichern unter","&Speichern", 1001, $path)
		 Send("!s")

		 $numPDFs = $numPDFs + 1
		 myLog($path & " wurde gespeichert", "202")
	  Else
		 $k = $k + 1
	  EndIf
   Until $k >= $max

   WinActivate("Gebührenfakturierung ","&Übermitteln")

   return $numPDFs

EndFunc


; Return the current time in the format YYYYMMDD_HHMMSS
Func getTime()
   Local $time = _NowCalc()
   $time = StringReplace($time, "/", "")
   $time = StringReplace($time, ":", "")
   $time = StringReplace($time, " ", "_")

   return $time
EndFunc

; Remove illegal chars for filepaths
Func removeIllegalChars($string)
   $string = StringReplace($string,"/","_")
   $string = StringReplace($string,":","_")
   $string = StringReplace($string," ","_")
   $string = StringReplace($string,"\","_")
   $string = StringReplace($string,".","_")
   $string = StringReplace($string,"+","_")
   $string = StringReplace($string,";","_")
   $string = StringReplace($string,"%","_")
   $string = StringReplace($string,"&","_")
   $string = StringReplace($string,"?","_")
   $string = StringReplace($string,"@","_")
   $string = StringReplace($string,"!","_")
   $string = StringReplace($string,"<","_")
   $string = StringReplace($string,">","_")
   $string = StringReplace($string,"*","_")
   $string = StringReplace($string,"|","_")
   $string = StringReplace($string,'"',"_")

   return $string
EndFunc

; Log into a .csv file
Func myLog($message, $id = "0", $type = "info", $logFile = $logFile)
   $type = StringUpper($type)
   $type = "["&$type&"]"

   While StringLen($id) < 4
	  $id = "0" & $id
   WEnd

   Local $time = _NowCalc()
   Local $log = $time & ";" & $id & ";" & $type & ";" & $message

   FileWriteLine($logFile, $log)

   return True
EndFunc

; Get window class name
Func _WinGetClass($hWnd)
    If IsHWnd($hWnd) = 0 And WinExists($hWnd) Then $hWnd = WinGetHandle($hWnd)
    Local $aGCNDLL = DllCall('User32.dll', 'int', 'GetClassName', 'hwnd', $hWnd, 'str', '', 'int', 4095)
    If @error = 0 Then Return $aGCNDLL[2]
    Return SetError(1, 0, '')
 EndFunc


;   __    __    ___  _        __   ___   ___ ___    ___      ______   ___       ___    ____   ____  _       ___    ____
; |  |__|  |  /  _]| |      /  ] /   \ |   |   |  /  _]    |      | /   \     |   \  |    | /    || |     /   \  /    |
; |  |  |  | /  [_ | |     /  / |     || _   _ | /  [_     |      ||     |    |    \  |  | |  o  || |    |     ||   __|
; |  |  |  ||    _]| |___ /  /  |  O  ||  \_/  ||    _]    |_|  |_||  O  |    |  D  | |  | |     || |___ |  O  ||  |  |
; |  `  '  ||   [_ |     /   \_ |     ||   |   ||   [_       |  |  |     |    |     | |  | |  _  ||     ||     ||  |_ |
;  \      / |     ||     \     ||     ||   |   ||     |      |  |  |     |    |     | |  | |  |  ||     ||     ||     |
;   \_/\_/  |_____||_____|\____| \___/ |___|___||_____|      |__|   \___/     |_____||____||__|__||_____| \___/ |___,_|
;
; Quote: GemoWin NG ist eine durchgängige Gesamtlösung für öffentliche Verwaltungen. Im Vordergrund von GemoWin NG steht
;        die Integration der Geschäftsprozesse und die zentrale Datenbewirtschaftung. Daten müssen nur einmal erfasst
;        werden und stehen allen Modulen zur Verfügung, dies erhöht die Qualität und die Effizienz.
;        (Source: https://www.dialog.ch/de/produkte/gemowin-ng/)
;
;        Allright, but exporting your data SUCKS HARD!