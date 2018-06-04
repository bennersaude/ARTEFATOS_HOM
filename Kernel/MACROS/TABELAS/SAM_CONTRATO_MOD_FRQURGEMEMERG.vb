'HASH: 769056E2A680A654B354145AA9762DF9
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Condicao As String

  Set Interface = NewQuery

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAINICIAL").Value >CurrentQuery.FieldByName("DATAFINAL").Value Then
      bsShowMessage("A Data Inicial não pode ser maior que a Data Final", "E")
      CanContinue = False
    End If
  End If


  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "AND CONTRATOMOD     = " + CStr(RecordHandleOfTable("SAM_CONTRATO_MOD"))

  Condicao = " CONTRATOMOD     = " + CStr(RecordHandleOfTable("SAM_CONTRATO_MOD"))

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_MOD_FRQURGEMEMERG", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime,"" , Condicao)

  If Linha <>"" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  Set Interface = Nothing

End Sub
