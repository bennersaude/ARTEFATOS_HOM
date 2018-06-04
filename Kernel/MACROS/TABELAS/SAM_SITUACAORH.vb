'HASH: E8ECB76CB40F4429D294E2BBDDCAB450
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("TABTIPOSITUACAO").AsInteger = 5) And (CurrentQuery.FieldByName("MIGRACAOEQUIVALENCIAMODULO").AsInteger = 2) Then
    If (CurrentQuery.FieldByName("FALECIMENTO").AsString = "S") Then
      BsShowMessage("Para migrar por equivalência de módulos o parâmetro 'Falecimento' não pode estar marcado","E")
      CanContinue = False
      Exit Sub
    End If
    'CurrentQuery.FieldByName("CONSIDMODOPCIONAIS").AsString =  "S"
  End If

  'sms 60045
  If (CurrentQuery.FieldByName("TABTIPOSITUACAO").AsInteger = 5) And (CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull) Then
    BsShowMessage("Motivo cancelamento deve ser informado","E")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("TABTIPOSITUACAO").AsInteger = 2) Then
    If (CurrentQuery.FieldByName("DIASCANCFUTURO").AsInteger > 0) And (CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull) Then
      BsShowMessage("Motivo cancelamento deve ser informado","E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("DIASCANCFUTURO").AsInteger = 0) And (Not CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull) Then
      BsShowMessage("Dias para cancelamento futuro deve ser informado","E")
      CanContinue = False
      Exit Sub
    End If

  End If
End Sub

