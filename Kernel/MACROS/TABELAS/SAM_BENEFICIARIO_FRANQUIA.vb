'HASH: 2C2401C92C9E5B8BEF90D77E93E7933F
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = " AND DATAFINAL < DATAINICIAL"
  Condicao = " AND FRANQUIA = " + CurrentQuery.FieldByName("FRANQUIA").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_FRANQUIA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Set Interface = Nothing
    Exit Sub
  End If

End Sub

