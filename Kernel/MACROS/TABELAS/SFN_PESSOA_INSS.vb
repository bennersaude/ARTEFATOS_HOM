'HASH: 52262B9A7F7C9E73EA2E75346799EAFA
'Macro: SFN_PESSOA_INSS
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 1 Then
    CurrentQuery.FieldByName("CNPJEMPREGADOR").Clear
    CurrentQuery.FieldByName("EMPREGADOR").Clear
  End If

  If Not CurrentQuery.FieldByName("CNPJEMPREGADOR").IsNull Then
    If Not IsValidCGC(CurrentQuery.FieldByName("CNPJEMPREGADOR").AsString)Then
      CanContinue = False
      bsShowMessage("CNPJ Inválido!", "E")
      Exit Sub
    End If
  End If

  'Dim P As Object
  'Set P =NewQuery
  'P.Clear
  'P.Add("SELECT INSSFAIXAS FROM SFN_PARAMETROSFIN")
  'P.Active =True

  'If CurrentQuery.FieldByName("FAIXA").AsInteger >P.FieldByName("INSSFAIXAS").AsInteger Then
  '    CanContinue =False
  '    MsgBox "Número de faixa inválido! Válido somente até "+Str(P.FieldByName("INSSFAIXAS").AsInteger)
  '    Exit Sub
  'End If
  'P.Active =False

  Set P = Nothing
  'CHECA VIGENCIA
  Dim Interface As Object
  Dim Linha As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SFN_PESSOA_INSS", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "PESSOA", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

