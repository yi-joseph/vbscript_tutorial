' Hier ist ein Programm von If-Else-Statement mit Input von Kunden in VBScript
' Dateiname: if_else_statement_mit_input.vbs

' Vorbereitung: Variable deklarieren und initialisieren
Dim nameBenutzers, nummerBenutzers

' Hauptprogramm: 
' Grüßen
nameBenutzers = InputBox ("Wie heissen Sie? ", "Benutzereingabe", "Yi")
MsgBox "Servus, " & nameBenutzers & ", hier ist ein Programm von If-Else-Statement mit Input in VBscript! "

' Benutzereingabe fuer eine Nummer:
nummerBenutzers = InputBox("Geben Sie bitte eine Nummer ein: ", "Zahleneingabe", "0")
MsgBox "Sie haben die Nummer " & nummerBenutzers & " eingegeben. "

' If-Else-Statement:
if (nummerBenutzers > 3) Then
	MsgBox "Sie bekommen den If-Zweig, weil Ihre Nummer groesser als 3 ist. " 
Else 
	MsgBox "Sie bekommen den Else-Zweig, weil Ihre Nummer kleiner als 3 ist. "
end if

' Ausgabe: Verabschieden
MsgBox "Das war's, tschau!"
