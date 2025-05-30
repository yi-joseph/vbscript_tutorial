' Hier ist der Code für Referenzparameter
' Call by Reference
' Dateiname: referenzParameter.vbs

' Vorbereitung: Funktion mit Referenzparameter deklarieren
Function quadrat(ByRef b)
	b = b^2
	quadrat = b
end Function

' Hauptprogramm: 
dim x, ergebnis
x = 6

' Ausgabe:
MsgBox "Servus, hier ist die Funktion mit Referenzparameter! "
MsgBox "Der Wert von der Variable x ist: " & x

ergebnis = quadrat(x)
MsgBox "Das Ergebnis der Funktion ist: " & ergebnis
x = 8
MsgBox "Hier ist der geaenderte Wert der Variable x: " & x

ergebnis = quadrat(x)
MsgBox "Hier ist das geaenderte Ergebnis von der Funktion: " & ergebnis

MsgBox "Das was, tschuess! "


