' Hier ist ein Code für die While-Wend-Schleife
' Dateiname: while_wend_schleife.vbs
' Vorbereitung: Variable deklarieren und initialisieren
dim i, s
i = 0 : s = ""

' i = 0 fängt es von 1 an.
' i = 1 fängt es von 2 an. 
' ":" bedeutet, die beiden Anweisungen werden durch : in einer Zeile getrennt und nacheinander ausgeführt
' i = 1
' s = ""

' Hauptprogramm: 
MsgBox "Hello, hier ist eine While-Schleife in VBScript! "

while i < 6
	s = s & i & ", "
	MsgBox "So geht es los: " & i
	i = i + 1
	MsgBox s
Wend

' Ausgabe: 
MsgBox "Das war's, tschau! "
