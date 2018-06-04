'HASH: ABA44E13AB0FAF6E264C39E7FD32D498

'#Uses "*bsShowMessage"
'#Uses "*ProcuraGrauValido"

Public Sub SOLICITANTE_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
  vCriterio = "SOLICITANTE = 'S'"
  vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Prestador", False, "")

  If vHandle <>0 Then
    CurrentQuery.FieldByName("SOLICITANTE").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim busca As Object
  Set busca = NewQuery
  busca.Add("SELECT A.SOLICITANTE FROM CLI_ATENDIMENTO AT, CLI_AGENDA A")
  busca.Add("WHERE AT.HANDLE = :ATENDIMENTO AND AT.AGENDA = A.HANDLE")
  busca.ParamByName("ATENDIMENTO").AsInteger = RecordHandleOfTable("CLI_ATENDIMENTO")
  busca.Active = True

  CurrentQuery.FieldByName("SOLICITANTE").AsInteger = busca.FieldByName("SOLICITANTE").AsInteger
  Set busca = Nothing
End Sub

Public Sub BOTAOREVERTER_OnClick()
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.ReverterNegacao(CurrentSystem, CurrentQuery.FieldByName("EVENTO").AsInteger, _
                         CurrentUser, _
                         CurrentQuery.FieldByName("HANDLE").AsInteger, _
                         "X")
  Set AGENDA = Nothing
End Sub

Public Sub BOTAOREVERTEROUTROUSUARIO_OnClick()
  Dim OLESenha As Object
  Set OLESenha = CreateBennerObject("senha.rotinas")
  vUsuario = OLESenha.PegarUsuarioPadrao(CurrentSystem, CurrentUser, vUsuario)
  Set OLESenha = Nothing
  If vUsuario <0 Then
    Exit Sub
  End If
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.ReverterNegacao(CurrentSystem, CurrentQuery.FieldByName("EVENTO").AsInteger, _
                         vUsuario, _
                         CurrentQuery.FieldByName("HANDLE").AsInteger, _
                         "X")
  Set AGENDA = Nothing
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  If CurrentQuery.FieldByName("EVENTO").IsNull Then
    Exit Sub
  End If

  Dim vHandle As Long
  vHandle = ProcuraGrauValido(CurrentQuery.FieldByName("EVENTO").AsInteger, GRAU.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
End Sub

Public Sub TABLE_AfterCommitted()
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.GerarNegacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "X")
  Set AGENDA = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT A.RECURSO FROM CLI_ATENDIMENTO A WHERE A.HANDLE = :ATENDIMENTO")
  SQL.Add("AND EXISTS(SELECT 1 FROM CLI_RECURSO_USUARIO RU, CLI_RECURSO R WHERE R.HANDLE = A.RECURSO AND R.PRESTADOR = RU.PRESTADOR AND RU.USUARIO = :USUARIO)")
  SQL.ParamByName("ATENDIMENTO").AsInteger = RecordHandleOfTable("CLI_ATENDIMENTO")
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True
  If SQL.EOF Then
    CanContinue = False
    bsShowMessage("Usuário inválido!", "E")
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOREVERTER") Then
		BOTAOREVERTER_OnClick
	End If
	If (CommandID = "BOTAOREVERTEROUTROUSUARIO") Then
		BOTAOREVERTEROUTROUSUARIO_OnClick
	End If

End Sub
