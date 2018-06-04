'HASH: 86AA814BADC937871D9582C77F5800D7
'Macro: SAM_PRESTADOR_PROC_HABIL

'#Uses "*bsShowMessage"

' Mauricio Ibelli -sms 1946 -22/03/2001 -so grava prestador/evento/habilitacao nao existir na tabela sam_prestador_habilitacao ou o parecer for diferente
' Mauricio Ibelli -04/01/2002 -sms3165 -Se filial padrao do prestador for nulo não checar responsavel

Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Dim S As Object
  Set S = NewQuery
  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  'Garcia
  'If S.FieldByName("CONTROLEDEACESSO").Value ="N" Then
  '  Ok =True
  '  Set S=Nothing
  '  Exit Function
  'End If

  SQL.Add("Select SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.filialpadrao FROM SAM_PRESTADOR_PROC, sam_prestador WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.handle = SAM_PRESTADOR_PROC.prestador")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And((SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser)Or(SQL.FieldByName("FILIALPADRAO").IsNull)), True, False)
  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida." + Chr(13)
  End If
  If(SQL.FieldByName("RESPONSAVEL").AsInteger <>CurrentUser)And(Not(SQL.FieldByName("FILIALPADRAO").IsNull))Then
  Mensagem = Mensagem + "Usuário não é o responsável!"
End If
Set SQL = Nothing
End Function

Public Sub EVENTO_OnExit()
  If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT ULTIMONIVEL FROM SAM_TGE WHERE HANDLE = " + CurrentQuery.FieldByName("EVENTO").AsString)
    SQL.Active = True
    If SQL.FieldByName("ULTIMONIVEL").AsString <>"S" Then
      bsShowMessage("O evento deve ser último nível.", "I")
      CurrentQuery.FieldByName("EVENTO").Clear
      EVENTO.SetFocus
    End If
    Set SQL = Nothing
  End If
  CurrentQuery.FieldByName("HABILITACAO").Clear
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vHandle As Long

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_TGE.Z_DESCRICAO|SAM_TGE.NIVELAUTORIZACAO"
  vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'   and exists (select 1 from SAM_TGE_HABILITACAO where SAM_TGE.HANDLE=SAM_TGE_HABILITACAO.EVENTO)"
  vCampos = "Evento|Descrição|Nível"
  'vHandle =interface.Exec("SAM_TGE|SAM_TGE_HABILITACAO[SAM_TGE.HANDLE=SAM_TGE_HABILITACAO.EVENTO]",vColunas,2,vCampos,vCriterio,"Tabela Geral de Eventos",False,"")
  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
    CurrentQuery.FieldByName("HABILITACAO").Clear
  End If

  Set interface = Nothing

End Sub

Public Sub TABLE_AfterInsert()
  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_HABIL"
  End If
End Sub

Public Sub TABLE_AfterScroll()

	If WebMode Then
	 	EVENTO.WebLocalWhere= "SAM_TGE.ULTIMONIVEL = 'S'   and EXISTS (SELECT 1 FROM SAM_TGE_HABILITACAO WHERE SAM_TGE.HANDLE=SAM_TGE_HABILITACAO.EVENTO)"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  ' sms 1946
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT * FROM SAM_PRESTADOR_HABILITACAO A WHERE A.PRESTADOR = :PREST AND A.EVENTO = :EVENTO AND A.HABILITACAO = :HABIL")
  SQL.ParamByName("PREST").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.ParamByName("HABIL").Value = CurrentQuery.FieldByName("HABILITACAO").AsInteger
  SQL.Active = True
  If SQL.EOF Then
    If CurrentQuery.FieldByName("OPERACAO").AsString = "E" Then
      CanContinue = False
      bsShowMessage("Habilitação não econtrada. Operação incoerente", "E")
      Exit Sub
    End If
  End If

  If Not SQL.EOF Then
    If CurrentQuery.FieldByName("OPERACAO").Value = "I" Then
      If SQL.FieldByName("TEMPORARIO").Value = "S" Then
        If CurrentQuery.FieldByName("MOTIVOPARECER").IsNull Then
          'MsgBox "ok"
        Else
          bsShowMessage("Evento/Habilitação já cadastrada para o prestador.", "E")
          CanContinue = False
        End If
      Else
        If CurrentQuery.FieldByName("motivoparecer").IsNull Then
          bsShowMessage("Evento/Habilitação já cadastrada para o prestador.", "E")
          CanContinue = False
        Else
          'MsgBox "ok"
        End If
      End If
    End If
  End If

  SQL.Active = False
  Set SQL = Nothing
End Sub

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
  If WebMode Then
  	HABILITACAO.WebLocalWhere = "A.HANDLE IN " + "(Select H.HABILITACAO FROM SAM_TGE_HABILITACAO H WHERE H.EVENTO = @CAMPO(EVENTO))"
  ElseIf VisibleMode Then
  	HABILITACAO.LocalWhere = "SAM_HABILITACAO.HANDLE IN " + "(Select H.HABILITACAO FROM SAM_TGE_HABILITACAO H WHERE H.EVENTO = @EVENTO)"
  End If

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
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
  	HABILITACAO.WebLocalWhere = "A.HANDLE IN " + "(Select H.HABILITACAO FROM SAM_TGE_HABILITACAO H WHERE H.EVENTO = @CAMPO(EVENTO))"
  ElseIf VisibleMode Then
  	HABILITACAO.LocalWhere = "SAM_HABILITACAO.HANDLE IN " + "(Select H.HABILITACAO FROM SAM_TGE_HABILITACAO H WHERE H.EVENTO = @EVENTO)"
  End If

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

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser
End Sub
