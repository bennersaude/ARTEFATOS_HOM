'HASH: 304EAEAD763E3E9C3565153800BBCB93
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Balani SMS 49282 03/10/2005
  Dim interface As Object
  Dim Erro As String
  Set interface = CreateBennerObject("SAMGERAL.Vigencia")

  Erro = interface.Vigencia(CurrentSystem, "SAM_CONTRATO_AUXILIO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", "")

  If Erro <> "" Then
    CanContinue = False
    bsShowMessage(Erro, "E")
  End If

  Set interface = Nothing
  'final SMS 49282
End Sub
