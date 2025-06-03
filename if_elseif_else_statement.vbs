' Hier ist ein Programm von If-Elseif-Else-Statement in VBScript
' Dateiname: if_elseif_else_statement.vbs

' Vorbereitung: Variable deklarieren und initialieren
Dim loesung, nameBenutzers, nummerBenutzers
loesung = 3

' Hauptprogramm: 
' Grüßen
nameBenutzers = InputBox("Wie heissen Sie? ", "Benutzernamen eingeben", "Yi?")
MsgBox "Servus, " & nameBenutzers & "! Hier ist ein Programm von If-Elseif-Else-Statement in VBScript, viel Spass! "
' Benutzer eingabe für eine Nummer:
nummerBenutzers = InputBox ("Geben Sie bitte eine Nummer ein: ", "Eine Nummer eingeben", "0")

' InputBox gibt immer einen String zurück
' wenn man eine Integer bekommen will, muss man den String in Integer umwandeln mit CInt()
nummerBenutzers = CInt(nummerBenutzers)
MsgBox "Sie haben die Zahl: " & nummerBenutzers & " eingegeben, schauen wir Mal, ob Sie Recht haben ... "

' Das Statement: 
if (nummerBenutzers = loesung) Then
	MsgBox "Bingo, Sie haben Recht, die Loesung ist " & loesung & " ! "
elseif (nummerBenutzers > loesung) Then
	MsgBox "Leider haben Sie groessere Zahl eingegeben! "
Else
	MsgBox "Leider haben Sie kleinere Zahl eingegeben! "
end if

' Ausgabe: verabschieden
MsgBox "Das war's, schoenen Tag! "

