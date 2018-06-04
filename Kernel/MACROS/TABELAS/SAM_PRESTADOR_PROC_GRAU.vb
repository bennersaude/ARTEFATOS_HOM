'HASH: 4205A4283C4B9B4085F252F7E534A5C8
'#Uses "*bsShowMessage"

'Mauricio Ibelli - 12/12/2001 - smsxxxx - Novo processo para grau
'Mauricio Ibelli - 04/01/2002 - sms3165 - Se filial padrao do prestador for nulo não checar responsavel

Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Dim S As Object
  Set S = NewQuery
  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  SQL.Add("SELECT SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.FILIALPADRAO FROM SAM_PRESTADOR_PROC, SAM_PRESTADOR WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PROC.PRESTADOR")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And ((SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser) Or (SQL.FieldByName("FILIALPADRAO").IsNull)), True, False)

  Mensagem = ""

  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida." + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
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

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If

  '' ********************************************************************************
  '' Alterado em 18/02/2002 -- por Durval SMS 6353 -- Foi comentada esta verificação
  '' ********************************************************************************
  ''  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
  ''    MsgBox "Fase com data finalizada não pode ser alterada!"
  ''    CanContinue = False
  ''    Exit Sub
  ''  End If
  '' ********************************************************************************
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT * FROM SAM_PRESTADOR_GRAU A WHERE A.PRESTADOR = :PREST AND A.GRAU = :GRAU")
  SQL.ParamByName("PREST").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    If CurrentQuery.FieldByName("OPERACAO").AsString = "E" Then
      CanContinue = False
      bsShowMessage("Prestador/Grau não econtrado. Operação incoerente", "E")
      Exit Sub
    End If
  End If

  If Not SQL.EOF Then
    If CurrentQuery.FieldByName("OPERACAO").Value = "I" Then
      bsShowMessage("Grau já cadastrado para o prestador.", "E")
      CanContinue = False
    End If
  End If

  SQL.Active = False

  Set SQL = Nothing

  If CurrentQuery.State = 3 And _
     Not Ok Then
    CanContinue = False
    bsShowMessage(Mensagem, "E")
  End If

End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("RESPONSAVEL").AsInteger = CurrentUser

  If WebMode Then
    CurrentQuery.FieldByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
  End If
End Sub
