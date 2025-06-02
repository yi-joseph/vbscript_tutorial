' Hier ist ein Programm von For-Schleife mit Sub
' Dateiname: for_schleife_mit_sub.vbs

' Vorbereitung: sub definieren.
' Sub zum Vergroessern
sub vergroessern()
	MsgBox "Hier ist ein Programm zum Vergroessern! "
	Dim s, i
	s = ""
	for i = 1 to 5 step 1 
		MsgBox "Jetzt faengt es an: " & i
		s = s & i & ", "
		MsgBox s
	Next
end sub

sub verkleinern()
	Dim s, i
	s = ""
	MsgBox "Jetzt ist ein Programm zum verkleinern! "
	for i = 5 to 1 step -1 
		MsgBox "Jetzt faengt es an: " & i
		s = s & i & ", "
		MsgBox s
	Next
end sub

' Hauptprogramm: 

MsgBox "Servus, hier ist ein Programm von For-Schleife mit zwei Prozeduren! "
vergroessern()
verkleinern()

' Ausgabe
MsgBox "Das war's! "
MsgBox "Vielen Dank fuer Ihre Zeit! Tschau! "

