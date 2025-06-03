' Hier ist ein Beispielcode für Until-Schleife
' Until steht hinter "do".
' das heißt: die Bedingung von Until-Schleife wird vor jedem Schleifendurchlauf geprüft.
' Dateiname: until_schleife_1.vbs

' Vorbereitung: 
dim i, s
i = 1 : s = ""

' Hauptprogramm: 
do Until i >= 6
	s = s & i & ", "
	MsgBox "Jetzt faengt es an: " & i
	i = i + 1
	MsgBox s
Loop

' Ausgabe: 
MsgBox "Das war's, schoenen Tag! "


