'HASH: 32C57B8F8A94BA2210DA4BE07CF34F29
'Macro: ANS_TSS_PARAMETROS
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  NomeTabela   = "ANS_TSS_PARAMETROS"
  CampoData1   = "COMPETENCIAINICIAL"
  CampoData2   = "COMPETENCIAFINAL"
  DataInicial  = CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime
  DataFinal    = CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime
  Condicao     = ""
  HandleTabela = CurrentQuery.FieldByName("HANDLE").AsInteger

  Dim Interface As Object
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  msg = Interface.Vigencia(CurrentSystem, NomeTabela, CampoData1, CampoData2, DataInicial, DataFinal, "", Condicao, HandleTabela)

  Set Interface = Nothing

  If Len(msg) > 0 Then
    bsShowMessage(msg, "E")
    CanContinue = False
  End If

End Sub
