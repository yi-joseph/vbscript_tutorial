' Hier ist ein Programm von Select-Case-Statement in VBScript
' Dateiname: select_case_mittagsMenue.vbs

' Vorbereitung: Variable deklarieren und initialisieren
dim nameBenutzers, mittagsMenue, auswahlBenutzers

' Hauptprogramm: 
' Grüßen
nameBenutzers = InputBox ("Wie heissen Sie? ", "Benutzername eingeben", "Yi?")
MsgBox "Servus, " & nameBenutzers & "! Hier ist dein Menue zum Mittagsessen! "

' Benutzers Auswahle eingeben: 
' vbCrlf ist Zeilenumbruch im InputBox von VBScript
' "& _" bedeutet Zeichenketten-Verkettung über mehrere Zeilen
mittagsMenue = InputBox ("Was wollen Sie nehmen:" & vbCrlf & _ 
			"1: Burger" & vbCrlf & _
			"2: Doener" & vbCrlf & _
			"3: gegrillte Haxe" & vbCrlf & _
			"4: Salat" & vbCrlf & _
			"5: Pizza" & vbCrlf & _
			"6: Kuchen", "Bitte geben Sie eine Zahl fuer Mittagsessen ein!", "0")

' Benutzers Auswahl von String in Integer umwandeln: 
mittagsMenue = CInt(mittagsMenue)

Select Case mittagsMenue
	Case 1 
		auswahlBenutzers = "Burger"
	Case 2 
		auswahlBenutzers = "gegrillte Haxe"
	Case 3 
		auswahlBenutzers = "gegrillte Haxe"
	Case 4
		auswahlBenutzers = "Salat"
	Case 5 
		auswahlBenutzers = "Pizza"
	Case 6 
		auswahlBenutzers = "Kunchen"
	Case Else
		auswahlBenutzers = "nichts"
End Select

' Ausgabe: 
MsgBox "Sie haben " & auswahlBenutzers & " ausgewaehlt, guten Appetit! "
MsgBox "Bis zum naechsten Mal! "

