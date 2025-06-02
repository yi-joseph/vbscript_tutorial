' Hier ist eine Funktion von Addition
' Funktion gibt einen Wert zurück
' Aufpassen: der Wertname der Rückgabe der Funktion ist der gleiche Name von der Funktion. 
' Dateiname: funktion.vbs

' Funktion definieren: 
Function additionRechnen()
	additionRechnen = 5 + b
end Function

'Hauptprogramm: 
Dim a, b
b = 3
a = additionRechnen()

' Ausgabe: 
MsgBox "Servus, hier ist eine Funktion heißt additionRechnen! "
MsgBox "Das Ergebnis von Addition ist: " & a

