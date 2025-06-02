' Hier ist der Cope fuer for_next_schleife
' Initialisierung "i = 1"
' Zielwert: "to 5"
' Verhalten: "step 1" um 1 vergrößert
' Verhalten: "step -1" um 1 verkleinert

' Vorbereitung: 
dim i, s

s = ""

' Hauptprogramm: 
MsgBox "Servus, hier ist ein Programm von For-Schleife in VBScript!"

for i = 1 to 5 step 1 
	MsgBox "Jetzt faengt es an: " & i
	s = s & i & ", "
	MsgBox s
Next

' Ausgabe: 
MsgBox "Das war's, tschau!"


