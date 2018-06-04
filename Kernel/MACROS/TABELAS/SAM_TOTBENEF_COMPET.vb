'HASH: A8FA9E3284501E3E0A73F8509724E968
 
Public Sub BOTAOGERAR_OnClick()
  Dim OBJ As Object

  Set OBJ =CreateBennerObject("RESUMOBENEF.Rotinas")
  OBJ.TotalizaBenef(CurrentSystem,CurrentQuery.FieldByName("COMPETENCIA").AsDateTime,CurrentQuery.FieldByName("DATACALCULO").AsDateTime,CurrentQuery.FieldByName("SOBREESCREVER").AsString)
  Set OBJ =Nothing

End Sub
