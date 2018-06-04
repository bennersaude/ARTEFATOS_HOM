'HASH: A6DA4DFBBD93879F0E8E3D4BF76C51EC
'Macro: TV_FILTRO_ATE001

Public Sub TABLE_AfterScroll()

 If SessionVar("BENEFICIARIOATE001") <> "" Then
    CurrentQuery.FieldByName("BENEFICIARIO").AsString = SessionVar("BENEFICIARIOATE001")
    CurrentQuery.FieldByName("COMPETENCIA").AsString = SessionVar("DATASOLICTATE001")

    SessionVar("BENEFICIARIOATE001") = ""
    SessionVar("DATASOLICTATE001") = ""
 End If

End Sub
