' Hier ist ein Programm von For-Each-Next-Schleife
' Dateiname: for_each_next_schleife.vbs

' Vorbereitung: Ein Array a deklarieren 
Dim a(2), b, s
a(0) = "a"
a(1) = "b"
a(2) = "c"
s = ""

' Hauptprogramm und zwar For-Each-Next-Schleife
MsgBox "Servus, hier ist ein Programm von For-Each-Next-Schleife in VBScript"
for each b in a 
	s = s & b & ", "
	MsgBox "Jetzt faengt es an: " & b
	MsgBox s
Next

MsgBox "Das war's, tschuess! "

