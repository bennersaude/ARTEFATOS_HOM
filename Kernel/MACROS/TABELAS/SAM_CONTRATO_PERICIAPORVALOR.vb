'HASH: ACCF7BF5DE5B8EAAA74027A238AF4FE1
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object

  Condicao = " AND TIPOAUTORIZACAO = " + CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PERICIAPORVALOR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha <>"" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

End Sub

