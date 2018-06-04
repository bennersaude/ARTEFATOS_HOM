'HASH: 2226F224D7CB3952BBBD280F8C797BF0
'SAM_TGE_VIGENCIATUSS
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim interface As Object
  Dim Condicao As String

  Set interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString

  Linha = interface.Vigencia(CurrentSystem, "SAM_TGE_VIGENCIATUSS", _
          "DATAINICIAL", _
          "DATAFINAL", _
          CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
          CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
          "EVENTO", _
          Condicao, _
          0)

  If Linha <> "" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  Set interface = Nothing

End Sub
