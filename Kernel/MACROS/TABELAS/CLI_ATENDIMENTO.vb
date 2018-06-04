'HASH: BEF7E33C261F7991231FA8A6F7B8C2AE
'CLI_ATENDIMENTO
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAATENDIMENTO_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("Select A.RECURSO")
  SQL.Add("  FROM CLI_ATENDIMENTO A,")
  SQL.Add("       CLI_RECURSO R")
  SQL.Add(" WHERE A.HANDLE = :ATENDIMENTO")
  SQL.Add("   And A.RECURSO = R.HANDLE")
  SQL.Add("   And EXISTS(Select 1")
  SQL.Add("                FROM CLI_RECURSO_USUARIO RU")
  SQL.Add("               WHERE RU.PRESTADOR = R.PRESTADOR")
  SQL.Add("                 And RU.USUARIO = :USUARIO)")
  SQL.ParamByName("ATENDIMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True
  If SQL.EOF Then
    CanContinue = False
    bsShowMessage("Usuário inválido!", "E")
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    bsShowMessage("O atendimento já foi finalizado!","I")
    Exit Sub
  End If
  If CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
    bsShowMessage("O atendimento não foi iniciado!", "I")
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("O atendimento já foi cancelado!", "I")
    Exit Sub
  End If
  If Not InTransaction Then StartTransaction
  SQL.Active = False
  SQL.Clear
  SQL.Add("UPDATE CLI_ATENDIMENTO SET DATACANCELAMENTO = :DATA, HORACANCELAMENTO = :HORA WHERE HANDLE = :HANDLE")
  SQL.ParamByName("DATA").Value = ServerDate
  SQL.ParamByName("HORA").Value = TimeValue(Str(DatePart("h", ServerNow)) + ":" + Str(DatePart("n", ServerNow)))
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL
  If InTransaction Then Commit
  RefreshNodesWithTable("CLI_ATENDIMENTO")
  Set SQL = Nothing
End Sub

Public Sub BOTAOFINALIZAATENDIMENTO_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT A.RECURSO FROM CLI_ATENDIMENTO A")
  SQL.Add(" WHERE A.HANDLE = :ATENDIMENTO")
  SQL.Add("   AND EXISTS(SELECT 1 FROM CLI_RECURSO_USUARIO RU, CLI_RECURSO R")
  SQL.Add("               WHERE R.HANDLE = A.RECURSO ")
  SQL.Add("                 AND R.PRESTADOR = RU.PRESTADOR ")
  SQL.Add("                 AND RU.USUARIO = :USUARIO)")
  SQL.ParamByName("ATENDIMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True
  If SQL.EOF Then
    CanContinue = False
    bsShowMessage("Usuário inválido!", "E")
    Exit Sub
  End If
  If CurrentQuery.State = 1 Then
    'Verifica se foi cadastrado pelo menos um evento
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT 1 FROM CLI_EVENTOSREALIZADOS WHERE ATENDIMENTO = :ATENDIMENTO")
    SQL.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      'Verifica se foi iniciado o atendimento
      If Not CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
        'Verifica se existe alguma pendência EVENTO E EXAME
        SQL.Active = False
        SQL.Clear
        SQL.Add("SELECT N.HANDLE")
        SQL.Add("  FROM CLI_EXAMES_NEGACAO N,")
        SQL.Add("       CLI_EXAMES E")
        SQL.Add("WHERE E.ATENDIMENTO = :ATENDIMENTO")
        SQL.Add("   AND N.EXAME = E.HANDLE")
        SQL.Add("   AND N.SITUACAO = 'P'")
        SQL.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQL.Active = True
        If Not SQL.EOF Then
          bsShowMessage("Existem exames com negações pendentes!", "I")
          Exit Sub
        End If
        SQL.Active = False
        SQL.Clear
        SQL.Add("SELECT N.HANDLE")
        SQL.Add("  FROM CLI_EVENTOSREALIZADOS_NEGACAO N,")
        SQL.Add("       CLI_EVENTOSREALIZADOS E")
        SQL.Add("WHERE E.ATENDIMENTO = :ATENDIMENTO")
        SQL.Add("   AND N.EVENTOSREALIZADOS = E.HANDLE")
        SQL.Add("   AND N.SITUACAO = 'P'")
        SQL.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQL.Active = True
        If Not SQL.EOF Then
          bsShowMessage("Existem eventos com negações pendentes!", "I")
          Exit Sub
        End If

        Dim Parametro As Object
        Set Parametro = NewQuery

        Parametro.Clear
        Parametro.Add("SELECT * FROM CLI_PARAMETROGERAL")
        Parametro.Active = True

        If Parametro.FieldByName("CIDOBRIGATORIOATENDIMENTO").AsString = "S" Then
          'Verifica se foi cadastrado um CID principal
          SQL.Active = False
          SQL.Clear
          SQL.Add("SELECT CID ")
          SQL.Add("  FROM CLI_DIAGNOSTICO")
          SQL.Add(" WHERE ATENDIMENTO = :ATENDIMENTO")
          SQL.Add("   AND CIDPRINCIPAL = 'S'")
          SQL.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
          SQL.Active = True
          If SQL.EOF Then
            bsShowMessage("Não foi marcado nenhum CID como principal!", "I")
            Exit Sub
          End If
        End If

        Set Parametro = Nothing

        If Not InTransaction Then StartTransaction
        SQL.Active = False
        SQL.Clear
        SQL.Add("UPDATE CLI_ATENDIMENTO SET DATAFINAL = :DATA, HORAFINAL = :HORA,")
        SQL.Add("                           DATACANCELAMENTO = NULL, HORACANCELAMENTO = NULL WHERE HANDLE = :HANDLE")
        SQL.ParamByName("DATA").Value = ServerDate
        SQL.ParamByName("HORA").Value = TimeValue(Str(DatePart("h", ServerNow)) + ":" + Str(DatePart("n", ServerNow)))
        SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQL.ExecSQL
        'Verifica se esse atendimento é de retorno
        SQL.Active = False
        SQL.Clear
        SQL.Add("SELECT MAX(T.DATAFINAL) ULTIMADATA")
        SQL.Add("  FROM CLI_AGENDA A,")
        SQL.Add("       CLI_ATENDIMENTO T")
        SQL.Add(" WHERE A.MATRICULA = :MATRICULA")
        SQL.Add("   AND A.RECURSO = :RECURSO")
        SQL.Add("   AND A.ESPECIALIDADE = :ESPECIALIDADE")
        SQL.Add("   AND T.AGENDA = A.HANDLE")
        SQL.Add("   AND T.TABMEDICOCOLETIVO = 1")
        SQL.Add("   AND T.HANDLE <> :HANDLE")
        SQL.ParamByName("MATRICULA").AsInteger = CurrentQuery.FieldByName("MATRICULA").AsInteger
        SQL.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
        SQL.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
        SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQL.Active = True
        Dim vDataMarcada As Date
        Dim vUltimaData As Date
        Dim vDiasRetorno As Integer
        Dim qAGENDA As Object
        Set qAGENDA = NewQuery
        qAGENDA.Active = False
        qAGENDA.Add("SELECT DATAMARCADA FROM CLI_AGENDA WHERE HANDLE = :AGENDA")
        qAGENDA.ParamByName("AGENDA").AsInteger = RecordHandleOfTable("CLI_AGENDA")
        qAGENDA.Active = True
        Dim QPARAMETROS As Object
        Set QPARAMETROS = NewQuery
        QPARAMETROS.Active = False
        QPARAMETROS.Add("SELECT DIASRETORNO FROM CLI_PARAMETROGERAL")
        QPARAMETROS.Active = True
        If Not SQL.EOF Then
          vDataMarcada = qAGENDA.FieldByName("DATAMARCADA").AsDateTime
          vUltimaData = SQL.FieldByName("ULTIMADATA").AsDateTime
          vDiasRetorno = QPARAMETROS.FieldByName("DIASRETORNO").AsInteger
          If (vDataMarcada - vUltimaData > 0) And (vDataMarcada - vUltimaData < vDiasRetorno) Then
            SQL.Active = False
            SQL.Clear
            SQL.Add("UPDATE CLI_AGENDA SET RETORNO = :RETORNO WHERE HANDLE = :AGENDA")
            SQL.ParamByName("RETORNO").AsString = "S"
            SQL.ParamByName("AGENDA").AsInteger = RecordHandleOfTable("CLI_AGENDA")
            SQL.ExecSQL
          End If
        End If
        Set qAGENDA = Nothing
        Set QPARAMETROS = Nothing
        If InTransaction Then Commit
        RefreshNodesWithTable("CLI_ATENDIMENTO")
      Else
        bsShowMessage("É necessário dar o início do atendimento!", "I")
      End If
    Else
      bsShowMessage("É obrigatório gravar no mínimo um evento para cada atendimento!", "I")
    End If
    Set SQL = Nothing
  End If
End Sub

Public Sub BOTAOINICIAATENDIMENTO_OnClick()

  If CurrentQuery.State <> 1 Then
    Dim DIA, MES, ANO, HORA, MIN As String
    Dim AGENDA As Object
    Set AGENDA = NewQuery
    Dim ATENDIMENTO As Object
    Set ATENDIMENTO = NewQuery
    AGENDA.Clear
    AGENDA.Add("SELECT DATAMARCADA, HORAMARCADA, HORACHEGADA FROM CLI_AGENDA WHERE HANDLE = :AGENDA")
    AGENDA.ParamByName("AGENDA").Value = CurrentQuery.FieldByName("AGENDA").AsInteger
    AGENDA.Active = True
    If Not AGENDA.FieldByName("HORACHEGADA").IsNull Or CurrentQuery.FieldByName("TABMEDICOCOLETIVO").AsInteger = 2 Then
      ATENDIMENTO.Active = False
      ATENDIMENTO.Clear
      ATENDIMENTO.Add("SELECT A.DATAMARCADA,")
      ATENDIMENTO.Add("       A.HORAMARCADA")
      ATENDIMENTO.Add("  FROM CLI_ATENDIMENTO T, CLI_AGENDA A")
      ATENDIMENTO.Add(" WHERE T.RECURSO = :RECURSO")
      ATENDIMENTO.Add("   And T.HORAINICIAL IS NOT NULL")
      ATENDIMENTO.Add("   And T.HORAFINAL IS NULL")
      ATENDIMENTO.Add("   And T.HORACANCELAMENTO IS NULL")
      ATENDIMENTO.Add("   And T.AGENDA <> :AGENDA")
      ATENDIMENTO.Add("   And T.AGENDA = A.HANDLE")
      ATENDIMENTO.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
      ATENDIMENTO.ParamByName("AGENDA").AsInteger = CurrentQuery.FieldByName("AGENDA").AsInteger
      ATENDIMENTO.Active = True
      If ATENDIMENTO.EOF Then
        If CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
          CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = ServerDate
          CurrentQuery.FieldByName("HORAINICIAL").AsDateTime = TimeValue(Str(DatePart("h", ServerNow)) + ":" + Str(DatePart("n", ServerNow)))
        End If
      Else
        DIA = Trim(Str(DatePart("d", ATENDIMENTO.FieldByName("DATAMARCADA").AsDateTime)))
        If Len(DIA) = 1 Then
          DIA = "0" + DIA
        End If

        MES = Trim(Str(DatePart("m", ATENDIMENTO.FieldByName("DATAMARCADA").AsDateTime)))
        If Len(MES) = 1 Then
          MES = "0" + MES
        End If

        ANO = Trim(Str(DatePart("yyyy", ATENDIMENTO.FieldByName("DATAMARCADA").AsDateTime)))

        HORA = Trim(Str(DatePart("h", ATENDIMENTO.FieldByName("HORAMARCADA").AsDateTime)))
        If Len(HORA) = 1 Then
          HORA = "0" + HORA
        End If

        MIN = Trim(Str(DatePart("n", ATENDIMENTO.FieldByName("HORAMARCADA").AsDateTime)))
        If Len(MIN) = 1 Then
          MIN = "0" + MIN
        End If

        bsShowMessage("O agendamento marcado para " + _
               DIA + "/" + MES + "/" + ANO + " " + HORA + ":" + MIN + " " + _
               " foi iniciado e não foi finalizado!", "I")
      End If
    Else
      bsShowMessage("É necessário marcar a chegada do paciente!", "I")
    End If
    Set AGENDA = Nothing
    Set ATENDIMENTO = Nothing
  End If

End Sub

Public Sub TABLE_AfterInsert()
  Dim AGENDA As Object
  Set AGENDA = NewQuery
  AGENDA.Add("SELECT TABMEDICOCOLETIVO, TABTIPOPACIENTE FROM CLI_AGENDA WHERE HANDLE = :AGENDA")
  AGENDA.ParamByName("AGENDA").Value = CurrentQuery.FieldByName("AGENDA").AsInteger
  AGENDA.Active = True
  If Not AGENDA.EOF Then
    CurrentQuery.FieldByName("TABMEDICOCOLETIVO").AsInteger = AGENDA.FieldByName("TABMEDICOCOLETIVO").AsInteger
    If AGENDA.FieldByName("TABMEDICOCOLETIVO").AsInteger = 1 Then
      CurrentQuery.FieldByName("TABTIPOPACIENTE").AsInteger = AGENDA.FieldByName("TABTIPOPACIENTE").AsInteger
    End If
  End If
  Set AGENDA = Nothing
End Sub

Public Sub TABLE_AfterPost()
  Dim vSolicitante As Long
  Dim ATIVIDADE As Object
  Set ATIVIDADE = NewQuery
  'Busca o evento da agenda
  ATIVIDADE.Clear
  ATIVIDADE.Add("SELECT AG.EVENTO, AG.GRAU, AG.SOLICITANTE")
  ATIVIDADE.Add("  FROM CLI_AGENDA AG             ")
  ATIVIDADE.Add(" WHERE AG.HANDLE = :AGENDA       ")
  ATIVIDADE.ParamByName("AGENDA").Value = CurrentQuery.FieldByName("AGENDA").AsInteger
  ATIVIDADE.Active = True
  vSolicitante = ATIVIDADE.FieldByName("SOLICITANTE").AsInteger
  'Verifica se a atividade possui evento
  If Not ATIVIDADE.EOF Then
    Dim VERIFICAEVENTO As Object
    Set VERIFICAEVENTO = NewQuery
    VERIFICAEVENTO.Clear
    VERIFICAEVENTO.Add("SELECT EVENTO FROM CLI_EVENTOSREALIZADOS WHERE ATENDIMENTO = :ATENDIMENTO")
    VERIFICAEVENTO.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    VERIFICAEVENTO.Active = True
    'Verifica se já não foi cadastrado o evento da atividade
    If VERIFICAEVENTO.EOF And Not ATIVIDADE.FieldByName("EVENTO").IsNull And Not ATIVIDADE.FieldByName("GRAU").IsNull Then
      Dim EVENTO As Object
      Set EVENTO = NewQuery

      EVENTO.Add("INSERT INTO CLI_EVENTOSREALIZADOS  ")
      EVENTO.Add("       (HANDLE,                    ")
      EVENTO.Add("        ATENDIMENTO,               ")
      EVENTO.Add("        SOLICITANTE,               ")
      EVENTO.Add("        EVENTO,                    ")
      EVENTO.Add("        QUANTIDADE,                ")
      EVENTO.Add("        GRAU,                      ")
      EVENTO.Add("        TABMEDICOCOLETIVO,         ")
      EVENTO.Add("        TABTIPOPACIENTE,           ")
      EVENTO.Add("        BENEFICIARIO,              ")
      EVENTO.Add("        MATRICULA)                 ")
      EVENTO.Add(" VALUES                            ")
      EVENTO.Add("       (:HANDLE,                   ")
      EVENTO.Add("        :ATENDIMENTO,              ")
      EVENTO.Add("        :SOLICITANTE,              ")
      EVENTO.Add("        :EVENTO,                   ")
      EVENTO.Add("        :QUANTIDADE,               ")
      EVENTO.Add("        :GRAU,                     ")
      EVENTO.Add("        :TABMEDICOCOLETIVO,        ")
      EVENTO.Add("        :TABTIPOPACIENTE,          ")
      EVENTO.Add("        :BENEFICIARIO,             ")
      EVENTO.Add("        :MATRICULA)                ")
      EVENTO.ParamByName("EVENTO").DataType = ftInteger
      EVENTO.ParamByName("GRAU").DataType = ftInteger
      EVENTO.ParamByName("BENEFICIARIO").DataType = ftInteger
      EVENTO.ParamByName("MATRICULA").DataType = ftInteger
      EVENTO.ParamByName("HANDLE").Value = NewHandle("CLI_EVENTOSREALIZADOS")
      EVENTO.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      EVENTO.ParamByName("SOLICITANTE").Value = vSolicitante
      EVENTO.ParamByName("EVENTO").Value = ATIVIDADE.FieldByName("EVENTO").Value
      EVENTO.ParamByName("QUANTIDADE").Value = 1
      EVENTO.ParamByName("GRAU").Value = ATIVIDADE.FieldByName("GRAU").Value
      EVENTO.ParamByName("TABMEDICOCOLETIVO").Value = CurrentQuery.FieldByName("TABMEDICOCOLETIVO").AsInteger
      EVENTO.ParamByName("TABTIPOPACIENTE").Value = CurrentQuery.FieldByName("TABTIPOPACIENTE").AsInteger
      EVENTO.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").Value
      EVENTO.ParamByName("MATRICULA").Value = CurrentQuery.FieldByName("MATRICULA").Value
      EVENTO.ExecSQL

      Set EVENTO = Nothing
    End If
    Set VERIFICAEVENTO = Nothing
  End If
  Set ATIVIDADE = Nothing
  If CurrentQuery.FieldByName("TABMEDICOCOLETIVO").AsInteger = 2 Then
    Dim BUSCA As Object
    Dim INSERE As Object
    Dim APAGA As Object
    Set BUSCA = NewQuery
    Set INSERE = NewQuery
    Set APAGA = NewQuery
    BUSCA.Add("SELECT BENEFICIARIO,")
    BUSCA.Add("       MATRICULA,")
    BUSCA.Add("       TABTIPOPARTICIPANTE")
    BUSCA.Add("  FROM CLI_TURMA_PARTICIPANTE")
    BUSCA.Add(" WHERE TURMA = :TURMA")
    BUSCA.ParamByName("TURMA").Value = CurrentQuery.FieldByName("TURMA").AsInteger
    BUSCA.Active = True
    If Not BUSCA.EOF Then

      APAGA.Clear
      APAGA.Add("DELETE FROM CLI_PARTICIPANTE WHERE ATENDIMENTO = :ATENDIMENTO")
      APAGA.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      APAGA.ExecSQL
      While Not BUSCA.EOF
        INSERE.Clear
        INSERE.Add("INSERT INTO CLI_PARTICIPANTE     ")
        INSERE.Add("       (HANDLE,                  ")
        INSERE.Add("        TABTIPOPARTICIPANTE,     ")
        INSERE.Add("        BENEFICIARIO,            ")
        INSERE.Add("        MATRICULA,               ")
        INSERE.Add("        INCLUSAODATA,            ")
        INSERE.Add("        INCLUSAOUSUARIO,         ")
        INSERE.Add("        ATENDIMENTO,             ")
        INSERE.Add("        PARTICIPOU)              ")
        INSERE.Add("VALUES                           ")
        INSERE.Add("       (:HANDLE,                 ")
        INSERE.Add("        :TABTIPOPARTICIPANTE,    ")
        INSERE.Add("        :BENEFICIARIO,           ")
        INSERE.Add("        :MATRICULA,              ")
        INSERE.Add("        :INCLUSAODATA,           ")
        INSERE.Add("        :INCLUSAOUSUARIO,        ")
        INSERE.Add("        :ATENDIMENTO,            ")
        INSERE.Add("        :PARTICIPOU)             ")
        INSERE.ParamByName("BENEFICIARIO").DataType = ftInteger
        INSERE.ParamByName("MATRICULA").DataType = ftInteger
        INSERE.ParamByName("HANDLE").Value = NewHandle("CLI_PARTICIPANTE")
        INSERE.ParamByName("TABTIPOPARTICIPANTE").Value = BUSCA.FieldByName("TABTIPOPARTICIPANTE").Value
        INSERE.ParamByName("BENEFICIARIO").Value = BUSCA.FieldByName("BENEFICIARIO").Value
        INSERE.ParamByName("MATRICULA").Value = BUSCA.FieldByName("MATRICULA").Value
        INSERE.ParamByName("INCLUSAODATA").Value = ServerDate
        INSERE.ParamByName("INCLUSAOUSUARIO").Value = CurrentUser
        INSERE.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("HANDLE").Value
        INSERE.ParamByName("PARTICIPOU").Value = "N"
        INSERE.ExecSQL
        BUSCA.Next
      Wend
      
    End If
    Set BUSCA = Nothing
    Set INSERE = Nothing
    Set APAGA = Nothing
  End If
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT 1 FROM CLI_RECURSO_USUARIO RU, CLI_RECURSO R")
  SQL.Add(" WHERE R.HANDLE = :RECURSO")
  SQL.Add("   AND RU.USUARIO = :USUARIO")
  SQL.Add("   AND RU.PRESTADOR = R.PRESTADOR")
  SQL.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True
  If SQL.EOF Then
    CanContinue = False
    bsShowMessage("Usuário inválido!", "E")
    Exit Sub
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOCANCELAATENDIMENTO") Then
		BOTAOCANCELAATENDIMENTO_OnClick
	End If
	If (CommandID = "BOTAOFINALIZAATENDIMENTO") Then
		BOTAOFINALIZAATENDIMENTO_OnClick
	End If
	If (CommandID = "BOTAOINICIAATENDIMENTO") Then
		BOTAOINICIAATENDIMENTO_OnClick
	End If

End Sub
