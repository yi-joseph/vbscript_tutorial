' Hier ist ein Beispielcode für While-Schleife
' while steht hinter "do".
' das heißt: die Bedingung vor While-Schleife wird vor jedem Schleifendurchlauf geprüft.
' Dateiname: while_schleife_1.vbs

' Vorbereitung: 
dim i, s
i = 1 : s = ""

' Hauptprogramm: 
do while i <= 6
	s = s & i & ", "
	MsgBox "Jetzt faengt es an: " & i
	i = i + 1
	MsgBox s
Loop

' Ausgabe: 
MsgBox "Das war's, schoenen Tag! "


