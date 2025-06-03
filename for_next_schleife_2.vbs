' Hier ist der Cope fuer for_next_schleife
' Initialisierung "i = 1"
' Zielwert: "to 5"
' Verhalten: "step 1" um 1 vergrößert
' Verhalten: "step -1" um 1 verkleinert
' Aufpassen: eine Variable muss in for-Schleife erstmals initialisiert werden. 
' Dateiname: for_next_schleife_2.vbs

' Vorbereitung: 
dim i, s

s = ""

' Hauptprogramm: 
MsgBox "Servus, hier ist ein Programm von For-Schleife in VBScript! "

for i = 5 to 1 step -1 
	MsgBox "Jetzt faengt es an: " & i
	s = s & i & ", "
	MsgBox s
Next

' Ausgabe: 
MsgBox "Das war's, tschau!"
