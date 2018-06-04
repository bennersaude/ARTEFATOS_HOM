'HASH: F1556847AE6CB6A14AA6C97DD9C37CE3
'#Uses "*bsShowMessage"
Public Function CHECARTETOREEMBOLSO()
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object

  CHECARTETOREEMBOLSO = True

  Condicao = "AND TETOREEMBOLSO = " + CurrentQuery.FieldByName("TETOREEMBOLSO").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_TETOREEMBOLSO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha = "" Then
    CHECARTETOREEMBOLSO = False
  Else
    CHECARTETOREEMBOLSO = True
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  If CHECARTETOREEMBOLSO Then
    CanContinue = False
    RefreshNodesWithTable("SAM_BENEFICIARIO_TETOREEMBOLSO")
    Exit Sub
  End If

End Sub

