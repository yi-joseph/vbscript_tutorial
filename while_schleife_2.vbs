' Hier ist ein Beispielcode fuer While-Schleife
' while steht hinter "loop"
' Das heißt: die Bedingung wird nach jedem Schleifdurchlauf geprüft

' Vorbereitung: 
dim i, s
i = 1 : s = ""

' Hauptprogramm: 
do 
	s = s &  i & ", "
	MsgBox "Jetzt faengt es an: " & i
	i = i + 1
	MsgBox s
loop while i <= 6

