'HASH: C3E64F68BD9325794D25C171B1563003
'#Uses "*bsShowMessage"


Public Function CHECARTETOREEMBOLSO()
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object

  CHECARTETOREEMBOLSO = True

  Condicao = " AND DATAFINAL > =  DATAINICIAL "

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_FAMILIA_TETOREEMBOLSO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FAMILIA", Condicao)

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
    RefreshNodesWithTable("SAM_FAMILIA_TETOREEMBOLSO")
    Exit Sub
  End If

End Sub

