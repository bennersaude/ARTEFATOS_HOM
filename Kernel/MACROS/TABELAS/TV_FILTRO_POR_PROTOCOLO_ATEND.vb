'HASH: 0648E46F3337498C852572E29FF91035
Public Sub TABLE_AfterScroll()

 If SessionVar("NUMEROPROTOCOLOANS") <> "" Then
    CurrentQuery.FieldByName("NUMEROPROTOCOLOANS").AsString = SessionVar("NUMEROPROTOCOLOANS")
    SessionVar("NUMEROPROTOCOLOANS") = ""
 End If

End Sub
