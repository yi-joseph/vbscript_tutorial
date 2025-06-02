' Hier ist ein Beispielcode fuer Until-Schleife
' while steht hinter "loop"
' Das heißt: die Bedingung wird nach jedem Schleifdurchlauf geprüft
' Der Dateiname heißt until_schleife_2.vbs

' Vorbereitung: 
dim i, s
i = 1 : s = ""

' Hauptprogramm: 
do 
	s = s &  i & ", "
	MsgBox "Jetzt faengt es an: " & i
	i = i + 1
	MsgBox s
	if (i = 7) then 
		MsgBox "Der Wert 7 ist schon erreicht, tschau! "
		exit do
	end If
loop while i >= 2

' Ausgabe: 
MsgBox "Das war's, schoenen Tag! "
