'HASH: 9ECF039593EC01F5CE91C51CBC7ACF71
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("ANOCALENDARIO").AsDateTime < CDate("2010-01-01") Then
    CanContinue = False
    bsShowMessage("Não é permitido o processamento da Dmed para anos anteriores a 2010!", "E")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
    If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
      CanContinue = False
      bsShowMessage("Ao preencher a data inicial, a data final também deve ser preenchida!", "E")
      Exit Sub
    End If
  End If

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
      CanContinue = False
      bsShowMessage("Ao preencher a data final, a data inicial também deve ser preenchida!", "E")
      Exit Sub
    End If
  End If

  Dim Interface  As Object
  Dim viRetorno  As Integer
  Dim vsMensagem As String

  Set Interface = CreateBennerObject("BSDMED.Rotinas")
  viRetorno = Interface.Processar(CurrentSystem, CurrentCompany, CurrentQuery.FieldByName("ANOCALENDARIO").AsDateTime, CurrentQuery.FieldByName("TESOURARIA").AsInteger,  CurrentQuery.FieldByName("OPERADORA").AsInteger, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, vsMensagem)

  If viRetorno > 0 Then
    CanContinue = False
    Set Interface = Nothing
    bsShowMessage(vsMensagem, "E")
    Exit Sub
  Else
    Set Interface = Nothing
    bsShowMessage("Processo enviado para execução no servidor!", "I")
  End If

End Sub
