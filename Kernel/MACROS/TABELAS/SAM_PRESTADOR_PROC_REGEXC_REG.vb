'HASH: 049665CC8AFD62EC4ED10F9E25595ED2
'Macro: SAM_PRESTADOR_PROC_REGEXC_REG

'#Uses "*bsShowMessage"

Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DATAFINAL,RESPONSAVEL FROM SAM_PRESTADOR_PROC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser, True, False)
  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida" + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Public Sub TABLE_AfterInsert()
  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_REGEXC_REG"
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then 'Adicionado CurrentSystem na assinatura, pois dizia que os parametros estavam errado
																	'e não batiam com a assinatura - Ricardo Rocha 06/06/2007
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
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then 'Adicionado CurrentSystem na assinatura, pois dizia que os parametros estavam errado
																	'e não batiam com a assinatura - Ricardo Rocha 06/06/2007
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

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then 'Adicionado CurrentSystem na assinatura, pois dizia que os parametros estavam errado
																	'e não batiam com a assinatura - Ricardo Rocha 06/06/2007
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_REGEXC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_REGEXC")
  SQL.Active = True
  If SQL.FieldByName("OPERACAO").AsString = "E" Then
    CanContinue = False
    bsShowMessage("O tipo de operação não permite inserir registros nesta carga !", "E")
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT *                                          ")
  SQL.Add("  FROM SAM_PRESTADOR_PROC_REGEXC_REG              ")
  SQL.Add(" WHERE HANDLE <> :HANDLE                          ")
  SQL.Add("   AND REGIMEATENDIMENTO = :REGIMEATENDIMENTO     ")
  SQL.Add("   AND PRESTADORPROCREGEXC = :PRESTADORPROCREGEXC ")

  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.ParamByName("REGIMEATENDIMENTO").Value = CurrentQuery.FieldByName("REGIMEATENDIMENTO").Value
  SQL.ParamByName("PRESTADORPROCREGEXC").Value = CurrentQuery.FieldByName("PRESTADORPROCREGEXC").Value
  SQL.Active = True
  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Este regime já está cadastrado !", "E")
  End If
  Set SQL = Nothing
End Sub

