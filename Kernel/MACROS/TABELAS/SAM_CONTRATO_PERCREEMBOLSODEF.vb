'HASH: 0E207816D8CF7A718D8008C513D469BA
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim vcondicao As String


  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  vcondicao = " and CONTRATO=" + CurrentQuery.FieldByName("CONTRATO").AsString

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PERCREEMBOLSODEF", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PERCREEMBOLSO", vcondicao)

  If Linha <> "" Then
    bsShowMessage(Linha, "E")
    CanContinue = False
  End If

  Set Interface = Nothing

End Sub

