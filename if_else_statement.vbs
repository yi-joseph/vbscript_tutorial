' Hier ist ein Programm von If-Else-Statement in VBScript
' Dateiname: if_else_statement.vbs

' Vorbereitung: Variable deklarieren und initialisieren
Dim a, b
a = 2
b = 4

' Hauptprogramm: 
' Grüßen
MsgBox "Servus, hier ist ein Programm von If-Else-Statement in VBscript"

if (a > 3) Then
	MsgBox "Hier ist If-Zweig und der Wert von a: " & a
Else 
	MsgBox "Hier ist Else-Zweig und der Wert von b: " & b
end if

MsgBox "Das war's, tschau!"
