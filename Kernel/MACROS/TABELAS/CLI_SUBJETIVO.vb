'HASH: B59426525A2CA0F29DD2CC67DFB464D0
'MACRO: CLI_SUBJETIVO
'Version: 8.11.0.0
'#Uses "*bsShowMessage"

Dim vbInserirRegistro As Boolean

Option Explicit

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("DATAABERTURA").AsDateTime = ServerDate
  CurrentQuery.FieldByName("RECURSO").AsInteger       = CLng(SessionVar("CLI_RECURSO"))
  CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger = CLng(SessionVar("SAM_ESPECIALIDADE"))
  CurrentQuery.FieldByName("PACIENTE").AsInteger      = CLng(SessionVar("SAM_MATRICULA"))

  Dim HLocalAtendimento As Long
  HLocalAtendimento = SelecionarLocalAtendimento
  If HLocalAtendimento > 0 Then
    CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger = HLocalAtendimento
  End If

  Dim HTipoAtendimento As Long
  HTipoAtendimento = SelecionarTipoAtendimento
  If HTipoAtendimento > 0 Then
    CurrentQuery.FieldByName("TIPOATENDIMENTO").AsInteger = HTipoAtendimento
  End If

  Dim HPlano As Long
  HPlano = SelecionarPlano
  If HPlano > 0 Then
    CurrentQuery.FieldByName("PLANO").AsInteger = HPlano
  End If

  vbInserirRegistro = True


End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "GUARDARATENDIMENTOSESSAO" Then
    ArmazenarAtendimento
  End If

  If CommandID = "EXCLUIRATENDIMENTO" Then
    If CurrentQuery.FieldByName("DATAENCERRAMENTO").IsNull Then
      ExcluirAtendimento
    Else
      bsShowMessage("Ação não executada, pois o atendimento já foi encerrado (um atendimento finalizado não pode ser cancelado)!", "I")
    End If
  End If

  If CommandID = "SELECIONARPACIENTE" Then
    SessionVar("CLI_SUBJETIVO") = Str(CurrentQuery.FieldByName("HANDLE").AsInteger)
    SessionVar("SAM_MATRICULA") = Str(CurrentQuery.FieldByName("PACIENTE").AsInteger)
    SessionVar("SAM_PLANO") = Str(CurrentQuery.FieldByName("PLANO").AsInteger)

	AtualizarCabecalho

    bsShowMessage("Paciente selecionado com sucesso!", "I")
  End If

  If CommandID = "SELECIONARATENDIMENTO" Then
	AtualizarSessaoMatricula
  End If

  If CommandID = "ENCERRARATENDIMENTO" Then
	CanContinue = EncerrarAtendimento
  End If

  If CommandID = "IMPRIMIRENCAMINHAMENTOS" Then
  	CanContinue = ChamarRelatoriosDeImpressao("CLI026",3)

  End If

  If CommandID = "IMPRIMIRRECEITA" Then
  	SessionVar("CLI_SUBJETIVO") = CurrentQuery.FieldByName("HANDLE").AsString
  End If

  If CommandID = "IMPRIMIRPROCEDIMENTOS" Then
  	CanContinue = ChamarRelatoriosDeImpressao("CLI025",4)
  End If

End Sub

Public Function EncerrarAtendimento() As Boolean

  If CurrentQuery.FieldByName("DATAENCERRAMENTO").IsNull Then

    If Not PermiteEncerrarObito Then
      EncerrarAtendimento = False
      bsShowMessage("Para finalizar o registro de óbito é obrigatório o preenchimento de pelo mesmo um CID!", "E")
      Exit Function
    End If

	Dim sql As Object
	Dim possuiimpressao As Boolean

    Dim Msg As String

	possuiimpressao = False
    Set sql = NewQuery


    'Não permitir encerrar um atendimento se não houver pelo menos um problema cadastrado para o atendimento.

    If VerificarEspecialidade Then
      sql.Active = False
      sql.Clear
      sql.Add("SELECT COUNT(1) ACHOU               ")
      sql.Add("  FROM CLI_PACIENTEDIAGNOSTICO      ")
      sql.Add(" WHERE ATENDIMENTO = :pAtendimento  ")
      sql.ParamByName("pAtendimento").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      sql.Active = True

      If (sql.FieldByName("ACHOU").AsInteger = 0) Then
        EncerrarAtendimento = False
        bsShowMessage("Ação não executada. Atendimento não possui nenhum problema cadastrado!", "E")
        Exit Function
      End If

    End If

	CurrentQuery.Edit

	sql.Active = False
	sql.Clear
	sql.Add("SELECT HANDLE FROM CLI_ENCAMINHAMENTOPEP WHERE SUBJETIVO = :ATENDIMENTO")
	sql.ParamByName("ATENDIMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active = True

	If Not sql.FieldByName("HANDLE").IsNull Then
	  CurrentQuery.FieldByName("IMPRIMIRENCAMINHAMENTO").AsString = "S"
	  possuiimpressao = True
	End If

	sql.Active = False
	sql.Clear
	sql.Add("SELECT HANDLE FROM CLI_PRESCRICAO PRE ")
	sql.Add(" WHERE ATENDIMENTO = :ATENDIMENTO     ")
	sql.Add("   AND EXISTS(SELECT HANDLE FROM CLI_PRESCRICAO_ITEM WHERE PRESCRICAO = PRE.HANDLE) ")
	sql.ParamByName("ATENDIMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active = True

	If Not sql.FieldByName("HANDLE").IsNull Then
	  CurrentQuery.FieldByName("IMPRIMIRRECEITA").AsString = "S"
	  possuiimpressao = True
	End If

	sql.Active = False
	sql.Clear
	sql.Add("SELECT PLA.HANDLE FROM CLI_PLANO_EXAMES PLA")
	sql.Add(" WHERE SUBJETIVO = :ATENDIMENTO            ")
	sql.Add("   AND EXISTS(SELECT HANDLE FROM CLI_EXAME_RESULTADO EXA WHERE EXAME = PLA.HANDLE)")

	sql.ParamByName("ATENDIMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active = True

	If Not sql.FieldByName("HANDLE").IsNull Then
	  CurrentQuery.FieldByName("IMPRIMIRPROCEDIMENTO").AsString = "S"
	  possuiimpressao = True
	End If

	sql.Active = False
	Set sql = Nothing

	CurrentQuery.FieldByName("DATAENCERRAMENTO").AsDateTime  = ServerDate
	CurrentQuery.FieldByName("DATAHORAALTERACAO").AsDateTime = ServerDate
	CurrentQuery.FieldByName("USUARIOALTERACAO").AsInteger  = CurrentUser
	CurrentQuery.Post

	Msg = "Ação executada, o atendimento foi finalizado com sucesso."
	If possuiimpressao = False Then
		Msg = Msg + " Não foi encontrado nenhum registro de impressão."
	End If

	bsShowMessage(Msg, "I")
  Else
  	EncerrarAtendimento = False
  	bsShowMessage("Ação não executada, pois o atendimento já havia sido finalizado!", "E")
   Exit Function
  End If

  EncerrarAtendimento = True
  DesligaPaciente

End Function


Public Function ChamarRelatoriosDeImpressao(vCodigoRelatorio As String, vOrigemRelatorio As Integer) As Boolean

	Dim sql As Object
	Dim qInsert As Object

	Set sql = NewQuery

	sql.Active = False
	sql.Clear
	sql.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = :CODIGO")
	sql.ParamByName("CODIGO").AsString = vCodigoRelatorio
	sql.Active = True

	If Not sql.FieldByName("HANDLE").IsNull Then
		Set qInsert = NewQuery
		qInsert.Active = False
		qInsert.Clear
		qInsert.Add("INSERT INTO CLI_AUDITORIA_RELATORIO")
	    qInsert.Add("       (HANDLE, USUARIO, DATAHORA, RELATORIO)")
	    qInsert.Add("VALUES (:HANDLE, :USUARIO, :DATAHORA, :RELATORIO)")
		qInsert.ParamByName("HANDLE").AsInteger = NewHandle("CLI_AUDITORIA_RELATORIO")
		qInsert.ParamByName("USUARIO").AsInteger = CurrentUser
		qInsert.ParamByName("DATAHORA").AsDateTime = ServerNow
	    qInsert.ParamByName("RELATORIO").AsInteger = sql.FieldByName("HANDLE").AsInteger
	    qInsert.ExecSQL

		Set qInsert = Nothing

		SessionVar("HATENDIMENTO") = CurrentQuery.FieldByName("HANDLE").AsString
		ReportPreview(sql.FieldByName("HANDLE").AsInteger,"",False,False)
		Set sql = Nothing
		ChamarRelatoriosDeImpressao = True
		Exit Function
	Else
		ChamarRelatoriosDeImpressao = False
		bsshowmessage("Relatório não encontrado no sistema","E")
		Set sql = Nothing
		Exit Function

	End If

End Function


Public Sub ArmazenarAtendimento()
  Dim vsMensagem As String

  SessionVar("CLI_SUBJETIVO") = CurrentQuery.FieldByName("HANDLE").AsString
  vsMensagem = "O atendimento foi selecionado com sucesso"

  AtualizarCabecalho

  bsShowMessage(vsMensagem, "I")
End Sub

Public Sub AtualizarCabecalho
  Dim Mensagem As String
  Dim SPP1 As BStoredProc

  Set SPP1 = NewStoredProc
  SPP1.Name = "BS_52316F90"  'benner.sc.clinica.cabecalho
  SPP1.AutoMode = True

  SPP1.AddParam("p_HMatricula",ptInput,ftInteger)
  SPP1.AddParam("p_HAtendimento",ptInput,ftInteger)
  SPP1.AddParam("p_HUsuario",ptInput,ftInteger)

  SPP1.ParamByName("p_HMatricula").AsInteger = CurrentQuery.FieldByName("PACIENTE").AsInteger
  SPP1.ParamByName("p_HAtendimento").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SPP1.ParamByName("p_HUsuario").AsInteger = CurrentUser

  SPP1.ExecProc

  Set SPP1 = Nothing

End Sub

Public Sub AtualizarSessaoMatricula
  Dim Mensagem As String
  Dim SPP1 As BStoredProc

  Set SPP1 = NewStoredProc
  SPP1.Name = "BS_2A390FB0"  'benner.sc.clinica.AtualizaSessaoMatricula
  SPP1.AutoMode = True

  SPP1.AddParam("p_HUsuario",ptInput,ftInteger)
  SPP1.AddParam("p_HMatricula",ptInput,ftInteger)

  SPP1.ParamByName("p_HUsuario").AsInteger = CurrentUser
  SPP1.ParamByName("p_HMatricula").AsInteger = CLng(CurrentVirtualQuery.FieldByName("BENEFICIARIO").AsInteger)

  SPP1.ExecProc

  Set SPP1 = Nothing

  Dim q1 As Object
  Set q1 = NewQuery

  q1.Active = False
  q1.Clear
  q1.Add("SELECT MATRICULA, SAMPLANO, ATENDIMENTO ")
  q1.Add("  FROM CLI_USUARIO_SESSAO            ")
  q1.Add(" WHERE USUARIO = :USUARIO            ")
  q1.ParamByName("USUARIO").AsInteger = CurrentUser
  q1.Active = True

  SessionVar("CLI_SUBJETIVO") = Str(q1.FieldByName("ATENDIMENTO").AsInteger)
  SessionVar("SAM_MATRICULA") = Str(q1.FieldByName("MATRICULA").AsInteger)
  SessionVar("SAM_PLANO") = Str(q1.FieldByName("SAMPLANO").AsInteger)

  q1.Active = False
  Set q1 = Nothing

End Sub

Public Sub TABLE_AfterPost()

  If (vbInserirRegistro = True) Then
    SessionVar("CLI_SUBJETIVO") = Str(CurrentQuery.FieldByName("HANDLE").AsInteger)
  End If

End Sub

Public Sub ExcluirAtendimento()

Dim q1 As Object
Set q1 = NewQuery

Dim SUBHANDLE As Long

  SUBHANDLE = CurrentQuery.FieldByName("HANDLE").AsFloat

  q1.Active = False
  q1.Clear
  q1.Add("DELETE FROM CLI_ANTECEDENTES WHERE ATENDIMENTO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsFloat = CurrentQuery.FieldByName("HANDLE").AsFloat
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE FROM CLI_DIAGNOSTICO WHERE ATENDIMENTO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE FROM CLI_ENCAMINHAMENTO_PROBLEMA                 ")
  q1.Add("  WHERE ENCAMINHAMENTO IN (SELECT HANDLE                 ")
  q1.Add("                             FROM CLI_ENCAMINHAMENTOPEP  ")
  q1.Add("                            WHERE SUBJETIVO = :SUBHANDLE)")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_ENCAMINHAMENTOPEP WHERE SUBJETIVO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_EXAME_FISICO WHERE SUBJETIVO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_EXAME_RESULTADO WHERE ATENDIMENTO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_PACIENTEDIAGNOSTICO WHERE ATENDIMENTO = :SUBHANDLE ")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_PACIENTEDIAGNOSTICO_HIST WHERE ATENDIMENTO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_EXAME_RESULTADO                      ")
  q1.Add("  WHERE EXAME In ( Select HANDLE                 ")
  q1.Add("                      FROM CLI_PLANO_EXAMES      ")
  q1.Add("                     WHERE SUBJETIVO = :SUBHANDLE)")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_PLANO_EXAMES_EXAME                           ")
  q1.Add("   WHERE EXAMESOLICIT IN ( SELECT HANDLE                 ")
  q1.Add("                             FROM CLI_PLANO_EXAMES       ")
  q1.Add("                            WHERE SUBJETIVO = :SUBHANDLE)")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_PLANO_EXAMES_PROBLEMA                           ")
  q1.Add("   WHERE EXAMESOLICIT IN ( SELECT HANDLE                    ")
  q1.Add("                             FROM CLI_PLANO_EXAMES          ")
  q1.Add("                            WHERE SUBJETIVO = :SUBHANDLE)   ")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_PLANO_EXAMES WHERE SUBJETIVO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_PLANO_OUTRASINDICACOES WHERE SUBJETIVO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

  q1.Active = False
  q1.Clear
  q1.Add(" DELETE CLI_PLANO_PROCPROPRIOS WHERE SUBJETIVO = :SUBHANDLE")
  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
  q1.ExecSQL

'  q1.Active = False
'  q1.Clear
'  q1.Add("  DELETE SAM_RELATORIOSWEBEMISSAO WHERE ATENDIMENTO = :SUBHANDLE ")
'  q1.ParamByName("SUBHANDLE").AsInteger = SUBHANDLE
'  q1.ExecSQL

  CurrentQuery.Delete
End Sub


Public Sub DesligaPaciente()
  Dim SQL As Object
  Set SQL = NewQuery

  If RegistradoObito Then
    Dim CIDObito As Long
    CIDObito = SelecionaCIDObito
    'Atualiza a matricula com a data e CID de falecimento
    SQL.Active = False
    SQL.Clear
    SQL.Add("UPDATE SAM_MATRICULA           ")
    SQL.Add("   SET DATAFALECIMENTO = :DATA,")
    SQL.Add("       CIDFALECIMENTO  = :CID  ")
    SQL.Add(" WHERE HANDLE = :HANDLE        ")
    SQL.ParamByName("DATA").AsDateTime  = CurrentQuery.FieldByName("DATAABERTURA").AsDateTime
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PACIENTE").AsInteger
    SQL.ParamByName("CID").AsInteger    = CIDObito
    SQL.ExecSQL
    Set SQL = Nothing
  End If
End Sub


Function SelecionaCIDObito As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT CID                     ")
  SQL.Add("  FROM CLI_PACIENTEDIAGNOSTICO ")
  SQL.Add(" WHERE ATENDIMENTO  = :HANDLE  ")
  SQL.Add("   AND EHCIDPRINCIPAL = 'S'    ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    SelecionaCIDObito = SQL.FieldByName("CID").AsInteger
  Else
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT CID                     ")
    SQL.Add("  FROM CLI_PACIENTEDIAGNOSTICO ")
    SQL.Add(" WHERE ATENDIMENTO  = :HANDLE  ")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    SelecionaCIDObito = SQL.FieldByName("CID").AsInteger
  End If
  SQL.Active = False
  Set SQL = Nothing
End Function


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If CLng(SessionVar("SAM_MATRICULA")) <= 0 Then
    CanContinue = False
    bsShowMessage("É necessário selecionar um paciente antes de abrir um novo atendimento!","E")
    Exit Sub
  End If

  If CLng(SessionVar("CLI_RECURSO")) <= 0 Then
    CanContinue = False
    bsShowMessage("O usuário não está vinculado a um profissional de saúde ativo para esta clínica!","E")
    Exit Sub
  End If

  If CLng(SessionVar("SAM_ESPECIALIDADE")) <= 0 Then
    CanContinue = False
    bsShowMessage("É necessário selecionar uma especialidade deste recurso antes de abrir um novo atendimento!","E")
    Exit Sub
  End If

  If TemConsultaAberta Then
    CanContinue = False
    bsShowMessage("É necessário finalizar o atendimento em aberto para este paciente antes de gerar um novo atendimento!","E")
    Exit Sub
  End If

  If TemRegistroObito Then
    CanContinue = False
    bsShowMessage("Foi registrado óbito para este participante!","E")
    Exit Sub
  End If

End Sub

Function TemRegistroObito As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT DATAFALECIMENTO  ")
  SQL.Add("  FROM SAM_MATRICULA    ")
  SQL.Add(" WHERE HANDLE = :HANDLE ")
  SQL.ParamByName("HANDLE").AsInteger = CLng(SessionVar("SAM_MATRICULA"))
  SQL.Active = True
  If Not SQL.FieldByName("DATAFALECIMENTO").IsNull Then
    TemRegistroObito = True
  Else
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT 1                                                       ")
    SQL.Add("  FROM CLI_SUBJETIVO       A                                   ")
    SQL.Add("  JOIN CLI_TIPOATENDIMENTO B ON (B.HANDLE = A.TIPOATENDIMENTO) ")
    SQL.Add(" WHERE A.PACIENTE = :PACIENTE                                  ")
    SQL.Add("   AND B.REGISTRODEOBITO = 'S'                                 ")
    SQL.Add("   AND A.HANDLE <> :HANDLE                                     ")
    SQL.ParamByName("PACIENTE").AsInteger = CLng(SessionVar("SAM_MATRICULA"))
    SQL.ParamByName("HANDLE").AsInteger   = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      TemRegistroObito = True
    Else
      TemRegistroObito = False
    End If
  End If
  Set SQL = Nothing
End Function

Function SelecionarLocalAtendimento As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE               ")
  SQL.Add("  FROM CLI_LOCALATENDIMENTO ")
  SQL.Add(" WHERE ASSUMIRPADRAO = 'S'  ")
  SQL.Active = True
  If Not SQL.EOF Then
    SelecionarLocalAtendimento = SQL.FieldByName("HANDLE").AsInteger
  Else
    SelecionarLocalAtendimento = -1
  End If
  Set SQL = Nothing
End Function

Function SelecionarPlano As Long
  If CLng(SessionVar("SAM_PLANO")) > 0 Then
    SelecionarPlano = CLng(SessionVar("SAM_PLANO"))
  Else

    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT C.PLANO                  ")
    SQL.Add("  FROM SAM_CONTRATO     C,      ")
    SQL.Add("       SAM_BENEFICIARIO B       ")
    SQL.Add(" WHERE B.CONTRATO  = C.HANDLE   ")
    SQL.Add("   AND B.MATRICULA = :MATRICULA ")
    SQL.Add("   AND ROWNUM     <= 1          ")
    SQL.ParamByName("MATRICULA").AsInteger = CLng(SessionVar("SAM_MATRICULA"))
    SQL.Active = True

    If Not SQL.EOF Then
      SelecionarPlano = SQL.FieldByName("PLANO").AsInteger
    Else
      SelecionarPlano = -1
    End If

    Set SQL = Nothing

  End If
End Function

Function SelecionarTipoAtendimento As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT T.TIPOATENDIMENTO                                                ")
  SQL.Add("  FROM CLI_TIPOATENDIMENTO_ESPEC T                                      ")
  SQL.Add("  JOIN CLI_USUARIO_SESSAO        U ON U.ESPECIALIDADE = T.ESPECIALIDADE ")
  SQL.Add(" WHERE T.ASSUMIRPADRAO = 'S'                                            ")
  SQL.Add("   AND U.USUARIO = :USUARIO                                             ")
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  If Not SQL.EOF Then
    SelecionarTipoAtendimento = SQL.FieldByName("TIPOATENDIMENTO").AsInteger
  Else
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT HANDLE                ")
    SQL.Add("  FROM CLI_TIPOATENDIMENTO   ")
    SQL.Add(" WHERE ASSUMIRPADRAO = 'S'   ")
    SQL.Active = True
    If Not SQL.EOF Then
      SelecionarTipoAtendimento = SQL.FieldByName("HANDLE").AsInteger
    Else
      SelecionarTipoAtendimento = -1
    End If
  End If
  Set SQL = Nothing
End Function

Function TemConsultaAberta As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                                                                          ")
  SQL.Add("  FROM CLI_SUBJETIVO S                                                            ")
  SQL.Add("  JOIN CLI_USUARIO_SESSAO U ON U.RECURSO = S.RECURSO AND U.MATRICULA = S.PACIENTE ")
  SQL.Add(" WHERE U.USUARIO  = :USUARIO                                                      ")
  SQL.Add("   AND S.DATAENCERRAMENTO IS NULL                                                 ")
  SQL.Add("   AND S.HANDLE <> :HANDLE                                                        ")
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    TemConsultaAberta = True
  Else
    TemConsultaAberta = False
  End If
  Set SQL = Nothing
End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim viDiasLimite As Integer
  If (Not DataDentroDoPeriodo(viDiasLimite)) And (CurrentQuery.State = 3) Then
    CanContinue = False
    CancelDescription = "A data especificada para a abertura do atendimento é anterior ao limite máximo permitido: " + Str(viDiasLimite) + " dias."
  End If

  If RegistradoObito Then
    If CurrentQuery.FieldByName("DATAABERTURA").AsDateTime > ServerDate Then
      CanContinue = False
      CancelDescription = "A data do atendimento não pode ser superior à data corrente, pois representa a data de falecimento do paciente!"
    ElseIf Not PrecisaInformarCID(True) Then
      RequestConfirmation(MontaMensagem)
    Else
      CanContinue = False
      CancelDescription = "Para finalizar o registro de óbito é obrigatório o preenchimento de pelo mesmo um CID!"
    End If
  End If
  If CurrentQuery.FieldByName("DATAABERTURA").AsDateTime > ServerDate Then
    CanContinue = False
    CancelDescription = "A data do atendimento não pode ser superior à data corrente!"
  End If

End Sub

Public Function PermiteEncerrarObito As Boolean
  PermiteEncerrarObito = True
  If RegistradoObito Then
    If PrecisaInformarCID(False) Then
      PermiteEncerrarObito = False
    End If
  End If
End Function

Function RegistradoObito As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                     ")
  SQL.Add("  FROM CLI_TIPOATENDIMENTO   ")
  SQL.Add(" WHERE REGISTRODEOBITO = 'S' ")
  SQL.Add("   AND HANDLE = :HANDLE      ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOATENDIMENTO").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    RegistradoObito = True
  Else
    RegistradoObito = False
  End If
  Set SQL = Nothing
End Function

Function PrecisaInformarCID(pbVerificarData As Boolean) As Boolean
  If (CurrentQuery.FieldByName("DATAENCERRAMENTO").IsNull And pbVerificarData) Then
    PrecisaInformarCID = False
  Else
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT 1                       ")
    SQL.Add("  FROM CLI_PACIENTEDIAGNOSTICO ")
    SQL.Add(" WHERE ATENDIMENTO = :HANDLE   ")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      PrecisaInformarCID = False
    Else
      PrecisaInformarCID = True
    End If
    SQL.Active = False
    Set SQL = Nothing
  End If
End Function

Function MontaMensagem As String
  Dim sAux As String
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT NOME                  ")
  SQL.Add("  FROM SAM_MATRICULA         ")
  SQL.Add(" WHERE HANDLE = :HANDLE      ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PACIENTE").AsInteger
  SQL.Active = True
  sAux = "O participante " + SQL.FieldByName("NOME").AsString + _
         " vai ser registrado como falecido e sua ficha médica só poderá ser consultada." + _
         " Deseja confirmar esta ação?"
  SQL.Active = False
  Set SQL = Nothing
  MontaMensagem = sAux
End Function

Function DataDentroDoPeriodo(piDiasLimite As Integer) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT DIASRETROATIVOSOAP ")
  SQL.Add("  FROM CLI_PARAMETROGERAL ")
  SQL.Active = True
  piDiasLimite = SQL.FieldByName("DIASRETROATIVOSOAP").AsInteger
  If Not SQL.EOF Then
    If (ServerDate - piDiasLimite) <= (CurrentQuery.FieldByName("DATAABERTURA").AsDateTime) Then
      DataDentroDoPeriodo = True
    Else
      DataDentroDoPeriodo = False
    End If
  Else
    DataDentroDoPeriodo = True
  End If
  Set SQL = Nothing
End Function


Function VerificarEspecialidade As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT OBRIGATORIOCID          ")
  SQL.Add("  FROM SAM_ESPECIALIDADE       ")
  SQL.Add(" WHERE HANDLE = :ESPECIALIDADE ")
  SQL.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  SQL.Active = True

  VerificarEspecialidade = (SQL.FieldByName("OBRIGATORIOCID").AsString = "S")

  Set SQL = Nothing
End Function

