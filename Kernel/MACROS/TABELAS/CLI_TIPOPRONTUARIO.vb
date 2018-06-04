'HASH: 0E4AA5AD970B1DED685F67C8401A9828


Public Sub BOTAOPREVER_OnClick()
  Dim PRONTUARIO As Object
  Set PRONTUARIO = CreateBennerObject("CliProntuario.Rotinas")
  PRONTUARIO.Consulta(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, -1,"", 0)
  Set PRONTUARIO = Nothing
End Sub

