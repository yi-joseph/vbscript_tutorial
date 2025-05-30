' Hier ist eine Funktion von Addition mit zwei Parameter
' Dateiname: parameter.vbs

'Vorbereitung: Funktion deklarieren: 
Function addition_zwei(a, b)
	addition_zwei = a + b
end Function

'Hauptprogramm: 
dim a, b
a = 3
b = 7

'Ausgabe: 
MsgBox "Servus, hier ist eine Addition von zwei Nummern"
MsgBox "3 + 7 ist ?"
MsgBox "Das Ergebnis ist: " & addition_zwei(a,b)
MsgBox "4 + 8 ist ? "
MsgBox "Das Ergebnis ist: " & addition_zwei(4,8)
