' Hier ist ein Beispielcode fuer Until-Schleife
' until steht hinter "loop"
' Das heißt: die Bedingung wird nach jedem Schleifdurchlauf geprüft
' Der Dateiname heißt until_schleife_2.vbs
' Until bedeutet: wenn eine Bedingung true ist, dann stopt das Programm. 

' Vorbereitung: 
dim i, s
i = 1 : s = ""

' Hauptprogramm: 
do 
	s = s &  i & ", "
	MsgBox "Jetzt faengt es an: " & i
	i = i + 1
	MsgBox s
loop Until i >= 5

' Ausgabe: 
MsgBox "Das war's, schoenen Tag! "
