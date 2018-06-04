'HASH: 168F4AB07541BF24D2347C31EAF9EFBE
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim vCondicao As String


  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  vCondicao = "AND CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PERCREEMBOLSO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PERCREEMBOLSO", vCondicao)

  If Linha <> "" Then
    bsShowMessage(Linha, "E")
    CanContinue = False
  End If

  Set Interface = Nothing


End Sub

