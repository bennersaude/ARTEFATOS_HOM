'HASH: 98AA8E7A150C21790EB47645A4CF6982
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q2 As Object
  Set q2 = NewQuery

  q2.Clear
  q2.Add("SELECT COUNT(1) QTD FROM SAM_TIPOCOMISSAO_INSCRICAO_PAR")
  q2.Add(" WHERE ORDEM = :pORDEM                ")
  q2.Add("   AND IDADEMAXIMA = :pIDADE          ")
  q2.Add("   AND TIPOCOMISSAOINSCRICAO = :pTIPO ")
  q2.ParamByName("pORDEM").AsInteger = CurrentQuery.FieldByName("ORDEM").AsInteger
  q2.ParamByName("pIDADE").AsInteger = CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger
  q2.ParamByName("pTIPO").AsInteger  = RecordHandleOfTable("SAM_TIPOCOMISSAO_INSCRICAO")
  q2.Active = True

  If q2.FieldByName("QTD").AsInteger > 0 Then
    MsgBox("Já existe um registro cadastrado para esta parcela e idade máxima!")
    CanContinue = False
    Exit Sub
  End If
  Set q1 = Nothing
  Set q2 = Nothing

End Sub
