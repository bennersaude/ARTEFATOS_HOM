'HASH: 59DBF0DE8B5AB3B42CB686A04C9AE24E
'Macro: SAM_MATRICULAHOMONIMA

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Obj As Object
  Set Obj = CreateBennerObject("SAM.GeraIniciais")
  CurrentQuery.FieldByName("INICIAIS").Value = Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("NOME").AsString)
  Set Obj = Nothing
End Sub

