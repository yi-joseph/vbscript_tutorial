' Hier ist der Werteparameter
' D.h. Call by Value. 
' Dateiname: werteParameter.vbs

' Vorbereitung: Funktion mit Werteparameter deklarieren: 
Function quadrat(ByVal a)
	a = a^2
	quadrat = a
end Function

' Hauptprogramm:
dim x, ergebnis
x = 4
ergebnis = quadrat(x)


' Ausgabe:
MsgBox "Servus, hier ist die Funktion mit Werteparameter! "
MsgBox "Der Wert von der Variable x ist 4"
MsgBox "Aber das Ergebnis der Funktion ist: " & ergebnis
MsgBox "Hier nochmal der Wert von der Variable x: " & x
MsgBox "Tschau! "
