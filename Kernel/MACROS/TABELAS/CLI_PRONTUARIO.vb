'HASH: 8AEB40ED28681953B2CA4C9D51998010
'CLI_PRONTUARIO
'#Uses "*bsShowMessage"

Public Sub BOTAOALTERAR_OnClick()
  Dim PRONTUARIO As Object
  If CurrentQuery.State <>1 Then
    bsShowMessage("Parâmetros em Edição!", "I")
    Exit Sub
  End If
  Dim USUARIO As Object
  Set USUARIO = NewQuery
  USUARIO.Clear
  USUARIO.Add("SELECT * FROM CLI_RECURSO_USUARIO WHERE USUARIO = :USUARIO")
  USUARIO.ParamByName("USUARIO").Value = CurrentUser
  USUARIO.Active = True
  If Not USUARIO.EOF Then
    Dim TIPOPRONTUARIO As Object
    Set TIPOPRONTUARIO = NewQuery
    TIPOPRONTUARIO.Clear
    TIPOPRONTUARIO.Add("SELECT VISUALIZACAODOPRONTUARIO FROM CLI_TIPOPRONTUARIO WHERE HANDLE = :TIPOPRONTUARIO")
    TIPOPRONTUARIO.ParamByName("TIPOPRONTUARIO").Value = CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger
    TIPOPRONTUARIO.Active = True
    If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "M" Then
      If USUARIO.FieldByName("PRONTUARIOMEDICO").AsString = "N" Then
        bsShowMessage("Usuário sem permissão!", "I")
        Exit Sub
      End If
    Else
      If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "S" Then
        If USUARIO.FieldByName("PRONTUARIOSOCIAL").AsString = "N" Then
          bsShowMessage("Usuário sem permissão!", "I")
          Exit Sub
        End If
      Else
        If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "O" Then
          If USUARIO.FieldByName("PRONTUARIOODONTOLOGICO").AsString = "N" Then
            bsShowMessage("Usuário sem permissão!", "I")
            Exit Sub
          End If
        End If
      End If
    End If
    Set TIPOPRONTUARIO = Nothing
    Set PRONTUARIO = CreateBennerObject("CliProntuario.Rotinas")
    PRONTUARIO.Altera(CurrentSystem, CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger, _
                      CurrentQuery.FieldByName("HANDLE").AsInteger, _
                      CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                      CurrentQuery.FieldByName("ATENDIMENTO").AsInteger, "", 0)
    Set PRONTUARIO = Nothing
  Else
    bsShowMessage("Usuário não está cadastrado no módulo de clínica!", "I")
  End If
  Set USUARIO = Nothing
End Sub

Public Sub BOTAOINSERIR_OnClick()
  Dim PRONTUARIO As Object
  Dim SQL As Object
  If CurrentQuery.State <>1 Then
    bsShowMessage("Parâmetros em Edição!", "I")
    Exit Sub
  End If
  Dim USUARIO As Object
  Set USUARIO = NewQuery
  USUARIO.Clear
  USUARIO.Add("SELECT * FROM CLI_RECURSO_USUARIO WHERE USUARIO = :USUARIO")
  USUARIO.ParamByName("USUARIO").Value = CurrentUser
  USUARIO.Active = True
  If Not USUARIO.EOF Then
    Dim TIPOPRONTUARIO As Object
    Set TIPOPRONTUARIO = NewQuery
    TIPOPRONTUARIO.Clear
    TIPOPRONTUARIO.Add("SELECT VISUALIZACAODOPRONTUARIO FROM CLI_TIPOPRONTUARIO WHERE HANDLE = :TIPOPRONTUARIO")
    TIPOPRONTUARIO.ParamByName("TIPOPRONTUARIO").Value = CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger
    TIPOPRONTUARIO.Active = True
    If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "M" Then
      If USUARIO.FieldByName("PRONTUARIOMEDICO").AsString = "N" Then
        bsShowMessage("Usuário sem permissão!", "I")
        Exit Sub
      End If
    Else
      If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "S" Then
        If USUARIO.FieldByName("PRONTUARIOSOCIAL").AsString = "N" Then
          bsShowMessage("Usuário sem permissão!", "I")
          Exit Sub
        End If
      Else
        If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "O" Then
          If USUARIO.FieldByName("PRONTUARIOODONTOLOGICO").AsString = "N" Then
            bsShowMessage("Usuário sem permissão!", "I")
            Exit Sub
          End If
        End If
      End If
    End If
    Set TIPOPRONTUARIO = Nothing
    Set PRONTUARIO = CreateBennerObject("CliProntuario.Rotinas")
    PRONTUARIO.Cadastro(CurrentSystem, CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger, _
                        CurrentQuery.FieldByName("ATENDIMENTO").AsInteger, _
                        CurrentQuery.FieldByName("HANDLE").AsInteger, 0, "")
    Set PRONTUARIO = Nothing
  Else
    bsShowMessage("Usuário não está cadastrado no módulo de clínica!", "I")
  End If
  Set USUARIO = Nothing
  RefreshNodesWithTable("CLI_PRONTUARIO")
End Sub

Public Sub BOTAOMOSTRAR_OnClick()
  Dim PRONTUARIO As Object
  If CurrentQuery.State <>1 Then
    bsShowMessage("Parâmetros em Edição!", "I")
    Exit Sub
  End If
  Dim USUARIO As Object
  Set USUARIO = NewQuery
  USUARIO.Clear
  USUARIO.Add("SELECT * FROM CLI_RECURSO_USUARIO WHERE USUARIO = :USUARIO")
  USUARIO.ParamByName("USUARIO").Value = CurrentUser
  USUARIO.Active = True
  If Not USUARIO.EOF Then
    Dim TIPOPRONTUARIO As Object
    Set TIPOPRONTUARIO = NewQuery
    TIPOPRONTUARIO.Clear
    TIPOPRONTUARIO.Add("SELECT VISUALIZACAODOPRONTUARIO FROM CLI_TIPOPRONTUARIO WHERE HANDLE = :TIPOPRONTUARIO")
    TIPOPRONTUARIO.ParamByName("TIPOPRONTUARIO").Value = CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger
    TIPOPRONTUARIO.Active = True
    If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "M" Then
      If USUARIO.FieldByName("PRONTUARIOMEDICO").AsString = "N" Then
        bsShowMessage("Usuário sem permissão!", "I")
        Exit Sub
      End If
    Else
      If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "S" Then
        If USUARIO.FieldByName("PRONTUARIOSOCIAL").AsString = "N" Then
          bsShowMessage("Usuário sem permissão!", "I")
          Exit Sub
        End If
      Else
        If TIPOPRONTUARIO.FieldByName("VISUALIZACAODOPRONTUARIO").AsString = "O" Then
          If USUARIO.FieldByName("PRONTUARIOODONTOLOGICO").AsString = "N" Then
            bsShowMessage("Usuário sem permissão!", "I")
            Exit Sub
          End If
        End If
      End If
    End If
    Set TIPOPRONTUARIO = Nothing
    Set PRONTUARIO = CreateBennerObject("CliProntuario.Rotinas")
    PRONTUARIO.Consulta(CurrentSystem, CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger, _
                        CurrentQuery.FieldByName("HANDLE").AsInteger, _
                        CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                        CurrentQuery.FieldByName("ATENDIMENTO").AsInteger, "", 0)
    Set PRONTUARIO = Nothing
  Else
    bsShowMessage("Usuário não está cadastrado no módulo de clínica!", "I")
  End If
  Set USUARIO = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim ATENDIMENTO As Object
  Set ATENDIMENTO = NewQuery
  ATENDIMENTO.Add("SELECT TABTIPOPACIENTE FROM CLI_ATENDIMENTO WHERE HANDLE = :ATENDIMENTO")
  ATENDIMENTO.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("ATENDIMENTO").AsInteger
  ATENDIMENTO.Active = True
  If Not ATENDIMENTO.EOF Then
    CurrentQuery.FieldByName("TABTIPOPACIENTE").AsInteger = ATENDIMENTO.FieldByName("TABTIPOPACIENTE").AsInteger
  End If
  Set AGENDA = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim vTemProntuario As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM CLI_PRONTUARIOCONTEUDO WHERE PRONTUARIO =:HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.FieldByName("HANDLE").IsNull Then
    vTemProntuario = True
  End If
  If Not vTemProntuario Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT MEMO FROM CLI_PRONTUARIO WHERE HANDLE = :PRONTUARIO")
    SQL.ParamByName("PRONTUARIO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If Not SQL.FieldByName("MEMO").IsNull Then
      vTemProntuario = True
    End If
  End If
  If vTemProntuario Then
    BOTAOINSERIR.Enabled = False
  Else
    BOTAOINSERIR.Enabled = True
  End If
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HORAFINAL, HORACANCELAMENTO FROM CLI_ATENDIMENTO WHERE HANDLE = :ATENDIMENTO")
  SQL.ParamByName("ATENDIMENTO").AsInteger = CurrentQuery.FieldByName("ATENDIMENTO").AsInteger
  SQL.Active = True
  If vTemProntuario Then
    BOTAOALTERAR.Enabled = True
    If(Not SQL.FieldByName("HORAFINAL").IsNull)Or(Not SQL.FieldByName("HORACANCELAMENTO").IsNull)Then
    BOTAOALTERAR.Enabled = False
    BOTAOINSERIR.Enabled = False
  End If
Else
  BOTAOALTERAR.Enabled = False
  If(Not SQL.FieldByName("HORAFINAL").IsNull)Or(Not SQL.FieldByName("HORACANCELAMENTO").IsNull)Then
  BOTAOINSERIR.Enabled = False
End If
End If
Set SQL = Nothing
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT HORAFINAL, HORACANCELAMENTO")
  SQL.Add("  FROM CLI_ATENDIMENTO")
  SQL.Add(" WHERE HANDLE = :ATENDIMENTO")
  SQL.ParamByName("ATENDIMENTO").AsInteger = RecordHandleOfTable("CLI_ATENDIMENTO")
  SQL.Active = True
  If(Not SQL.FieldByName("HORAFINAL").IsNull)Or(Not SQL.FieldByName("HORACANCELAMENTO").IsNull)Then
  CanContinue = False
  bsShowMessage("O atendimento já foi finalizado!", "Ï")
  Exit Sub
End If
Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT HORAFINAL, HORACANCELAMENTO")
  SQL.Add("  FROM CLI_ATENDIMENTO")
  SQL.Add(" WHERE HANDLE = :ATENDIMENTO")
  SQL.ParamByName("ATENDIMENTO").AsInteger = RecordHandleOfTable("CLI_ATENDIMENTO")
  SQL.Active = True
  If(Not SQL.FieldByName("HORAFINAL").IsNull)Or(Not SQL.FieldByName("HORACANCELAMENTO").IsNull)Then
  CanContinue = False
  bsShowMessage("O atendimento já foi finalizado!", "E")
  Exit Sub
End If
Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT A.RECURSO")
  SQL.Add("  FROM CLI_ATENDIMENTO A")
  SQL.Add(" WHERE A.HANDLE = :ATENDIMENTO")
  SQL.Add("   AND EXISTS(SELECT 1 FROM CLI_RECURSO_USUARIO RU, CLI_RECURSO R")
  SQL.Add("               WHERE R.HANDLE = A.RECURSO")
  SQL.Add("                 AND RU.PRESTADOR = R.PRESTADOR")
  SQL.Add("                 AND RU.USUARIO = :USUARIO)")
  SQL.ParamByName("ATENDIMENTO").AsInteger = CurrentQuery.FieldByName("ATENDIMENTO").AsInteger
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True
  If SQL.EOF Then
    CanContinue = False
    bsShowMessage("Usuário inválido!", "E")
    Exit Sub
  End If
  Set SQL = Nothing
  Dim BUSCA As Object
  Set BUSCA = NewQuery
  BUSCA.Add("SELECT P.HANDLE FROM CLI_PRONTUARIO P, CLI_TIPOPRONTUARIO T ")
  BUSCA.Add(" WHERE P.TIPOPRONTUARIO = :TIPOPRONTUARIO")
  BUSCA.Add("   AND P.ATENDIMENTO = :ATENDIMENTO")
  BUSCA.Add("   AND P.TIPOPRONTUARIO = T.HANDLE")
  BUSCA.Add("   AND T.PRONTUARIOUNICO = 'S'")
  BUSCA.ParamByName("TIPOPRONTUARIO").Value = CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger
  BUSCA.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("ATENDIMENTO").AsInteger
  BUSCA.Active = True
  If Not BUSCA.EOF Then
    bsShowMessage("Esse tipo de prontuário não pode ser repetido no mesmo atendimento!", "E")
    CanContinue = False
  End If
  Set BUSCA = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOALTERAR") Then
		BOTAOALTERAR_OnClick
	End If
	If (CommandID = "BOTAOINSERIR") Then
		BOTAOINSERIR_OnClick
	End If
	If (CommandID = "BOTAOMOSTRAR") Then
		BOTAOMOSTRAR_OnClick
	End If
End Sub
