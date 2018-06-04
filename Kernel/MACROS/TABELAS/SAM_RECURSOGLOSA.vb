'HASH: 44089E3F74FF1220964CF625D037CE05
'MACRO: SAM_RECURSOGLOSA


'SMS 76000 - Marcelo Barbosa - 29/01/2007
Public Sub BOTAOGERAR_OnClick()

  Dim vInterface As Object

  Set vInterface = CreateBennerObject("CA001.CATEND")

  vInterface.GeraLoteDef(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set vInterface = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  Dim ParamProcContas As Object

  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    Set ParamProcContas = NewQuery
    ParamProcContas.Active = False
    ParamProcContas.Clear
    ParamProcContas.Add("SELECT UTILIZADEFERIMENTO FROM SAM_PARAMETROSPROCCONTAS")
    ParamProcContas.Active = True

    If ParamProcContas.FieldByName("UTILIZADEFERIMENTO").AsString = "S" Then
      DATAPEDIDO.ReadOnly = False
    Else
      DATAPEDIDO.ReadOnly = True
    End If

    Set ParamProcContas = Nothing
  Else
    DATAPEDIDO.ReadOnly = True
  End If
End Sub
'FIM SMS 76000

