' Syntax für Prozeduren
' Prozeduren geben keinen Wert zurück. 
' prozeduren.vbs

' Vorbereitung für den Prozedur:
sub additionRechnen()
	a = 5 + b
	MsgBox "Das Ergebnis ist: " & a
end sub

' Hauptprogramm: 
dim a, b
b = 3

' Ausgabe: 
MsgBox "Servus!"
additionRechnen()


