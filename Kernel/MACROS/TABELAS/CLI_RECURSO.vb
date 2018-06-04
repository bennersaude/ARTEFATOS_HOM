'HASH: D6CB70D198E7E66E50C8674FE6CC7037
'#Uses "*bsShowMessage"
Option Explicit
Public Sub BOTAOAGENDA_OnClick()
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.Agendamento(CurrentSystem, CurrentQuery.FieldByName("CLINICA").AsInteger, _
                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                     0, 0, 0, 0)
  Set AGENDA = Nothing
End Sub

Public Sub BOTAOENTRADA_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição ou inserção!", "I")
    Exit Sub
  End If

  Dim vclinica, vdata As String
  Dim procura As Object
  Set procura = NewQuery
  procura.Add("SELECT HM.DATA, PC.NOME CLINICA")
  procura.Add("  FROM CLI_HORARIOMEDICO HM, CLI_CLINICA C, SAM_PRESTADOR PC")
  procura.Add(" WHERE RECURSO = :RECURSO AND HORASAIDA IS NULL")
  procura.Add("   AND HM.CLINICA = C.HANDLE AND C.PRESTADOR = PC.HANDLE")
  procura.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  procura.Active = True

  vdata = procura.FieldByName("DATA").AsString
  vclinica = procura.FieldByName("CLINICA").AsString

  If Not procura.EOF Then
    bsShowMessage("Existe um horário no dia " + vdata + " que está em aberto!", "I")
    Exit Sub
  End If
  Set procura = Nothing

  If Not InTransaction Then
    StartTransaction
  End If

  Dim insere As Object
  Set insere = NewQuery
  insere.Add("INSERT INTO CLI_HORARIOMEDICO")
  insere.Add("(HANDLE, CLINICA, RECURSO, DATA, HORAENTRADA, USUARIOENTRADA)")
  insere.Add("VALUES")
  insere.Add("(:HANDLE, :CLINICA, :RECURSO, :DATA, :HORAENTRADA, :USUARIOENTRADA)")
  insere.ParamByName("HANDLE").AsInteger = NewHandle("CLI_HORARIOMEDICO")
  insere.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
  insere.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  insere.ParamByName("DATA").AsDateTime = ServerDate
  insere.ParamByName("HORAENTRADA").AsDateTime = ServerNow
  insere.ParamByName("USUARIOENTRADA").AsInteger = CurrentUser
  insere.ExecSQL

  bsShowMessage("Marcado o horário de entrada!", "I")

  If InTransaction Then
    Commit
  End If

  Set insere = Nothing
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição ou inserção!", "I")
    Exit Sub
  End If

  Dim RelatorioHandle As Long
  Dim QueryBuscaHandle As Object
  Set QueryBuscaHandle = NewQuery
  QueryBuscaHandle.Active = False
  QueryBuscaHandle.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'CLI001'")
  QueryBuscaHandle.Active = True
  RelatorioHandle = QueryBuscaHandle.FieldByName("HANDLE").AsInteger

  On Error GoTo Final

  ReportPreview(RelatorioHandle, "CLI_RECURSO.HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString, True, False)
  Exit Sub

Final :
  bsShowMessage(Str(Error), "E")
End Sub

Public Sub BOTAOSAIDA_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição ou inserção!", "I")
    Exit Sub
  End If

  Dim procura As Object
  Set procura = NewQuery
  procura.Add("SELECT HM.HANDLE")
  procura.Add("  FROM CLI_HORARIOMEDICO HM")
  procura.Add(" WHERE RECURSO = :RECURSO")
  procura.Add("   AND CLINICA = :CLINICA ")
  procura.Add("   AND HORASAIDA IS NULL")
  procura.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  procura.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
  procura.Active = True

  If procura.EOF Then
    bsShowMessage("Não existe horário que está em aberto!", "I")
    Set procura = Nothing
    Exit Sub
  End If

  If Not InTransaction Then
    StartTransaction
  End If

  Dim altera As Object
  Set altera = NewQuery
  altera.Add("UPDATE CLI_HORARIOMEDICO SET HORASAIDA = :HORASAIDA, USUARIOSAIDA = :USUARIOSAIDA")
  altera.Add("WHERE HANDLE = :HANDLE")
  altera.ParamByName("HORASAIDA").AsDateTime = ServerNow
  altera.ParamByName("USUARIOSAIDA").AsInteger = CurrentUser
  altera.ParamByName("HANDLE").AsInteger = procura.FieldByName("HANDLE").AsInteger
  altera.ExecSQL

  bsShowMessage("Marcado o horário de saída!", "I")

  If InTransaction Then
    Commit
  End If

  Set altera = Nothing
  Set procura = Nothing
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vOrdemBusca As Integer
  Dim vbAux As Boolean


  vbAux = True
  On Error GoTo caracteres
  CDbl(PRESTADOR.Text)
  vOrdemBusca = 1
  vbAux = False
caracteres:
  If vbAux Then
    vOrdemBusca = 2
  End If

  ShowPopup = False

  Dim qBuscaPrestador As Object
  Dim Interface As Object
  Set qBuscaPrestador = NewQuery
  Set Interface = CreateBennerObject("Procura.Procurar")

  If PRESTADOR.Text <> "" Then 'Se nào foi digitado nada

    If vOrdemBusca = 2 Then 'o que foi digitado é o nome do prestador
      'Este select foi montado de acordo com o que já era executado na interface de procura
      qBuscaPrestador.Clear
      qBuscaPrestador.Add("SELECT COUNT(1) QTDE")
      qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
      qBuscaPrestador.Add(" WHERE HANDLE IN (SELECT PE.PRESTADOR ")
      qBuscaPrestador.Add("                    FROM CLI_CLINICA C, ")
      qBuscaPrestador.Add("                         SAM_PRESTADOR_PRESTADORDAENTID PE ")
      qBuscaPrestador.Add("                   WHERE C.PRESTADOR = PE.ENTIDADE ")
      qBuscaPrestador.Add("                     AND C.HANDLE = :HCLINICA ")
      qBuscaPrestador.Add("       )")
      qBuscaPrestador.Add("   AND UPPER(Z_NOME) LIKE ('" + PRESTADOR.Text + "%')")
      qBuscaPrestador.ParamByName("HCLINICA").AsInteger = RecordHandleOfTable("CLI_CLINICA")
      qBuscaPrestador.Active = True

      If (qBuscaPrestador.FieldByName("QTDE").AsInteger > 1) Or (qBuscaPrestador.FieldByName("QTDE").AsInteger = 0) Then 'A busca retornou mais de um registro, ou não retornou nenhum

        vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
        vCriterio = "HANDLE IN (SELECT PE.PRESTADOR FROM CLI_CLINICA C, SAM_PRESTADOR_PRESTADORDAENTID PE WHERE C.PRESTADOR = PE.ENTIDADE AND C.HANDLE = " + CurrentQuery.FieldByName("CLINICA").AsString + ")"

        vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
        vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Médico", False, PRESTADOR.Text)

      Else 'Se encontrou um prestador com o dado digitado

        qBuscaPrestador.Clear
        qBuscaPrestador.Add("SELECT HANDLE")
        qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
        qBuscaPrestador.Add(" WHERE HANDLE IN (SELECT PE.PRESTADOR ")
        qBuscaPrestador.Add("                    FROM CLI_CLINICA C, ")
        qBuscaPrestador.Add("                         SAM_PRESTADOR_PRESTADORDAENTID PE ")
        qBuscaPrestador.Add("                   WHERE C.PRESTADOR = PE.ENTIDADE ")
        qBuscaPrestador.Add("                     AND C.HANDLE = :HCLINICA ")
        qBuscaPrestador.Add("       )")
        qBuscaPrestador.Add("   AND UPPER(Z_NOME) LIKE ('" + PRESTADOR.Text + "%')")
        qBuscaPrestador.ParamByName("HCLINICA").AsInteger = RecordHandleOfTable("CLI_CLINICA")
        qBuscaPrestador.Active = True

        vHandle = qBuscaPrestador.FieldByName("HANDLE").AsInteger

      End If
    Else 'Foi digitado o código do prestador
      qBuscaPrestador.Clear
      qBuscaPrestador.Add("SELECT COUNT(1) QTDE")
      qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
      qBuscaPrestador.Add(" WHERE HANDLE IN (SELECT PE.PRESTADOR ")
      qBuscaPrestador.Add("                    FROM CLI_CLINICA C, ")
      qBuscaPrestador.Add("                         SAM_PRESTADOR_PRESTADORDAENTID PE ")
      qBuscaPrestador.Add("                   WHERE C.PRESTADOR = PE.ENTIDADE ")
      qBuscaPrestador.Add("                     AND C.HANDLE = :HCLINICA ")
      qBuscaPrestador.Add("       )")
      qBuscaPrestador.Add("   AND PRESTADOR LIKE '" + PRESTADOR.Text + "%'")
      qBuscaPrestador.ParamByName("HCLINICA").AsInteger = RecordHandleOfTable("CLI_CLINICA")
      qBuscaPrestador.Active = True

      If (qBuscaPrestador.FieldByName("QTDE").AsInteger > 1) Or (qBuscaPrestador.FieldByName("QTDE").AsInteger = 0) Then

        vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
        vCriterio = "HANDLE IN (SELECT PE.PRESTADOR FROM CLI_CLINICA C, SAM_PRESTADOR_PRESTADORDAENTID PE WHERE C.PRESTADOR = PE.ENTIDADE AND C.HANDLE = " + CurrentQuery.FieldByName("CLINICA").AsString + ")"

        vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
        vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, vCampos, vCriterio, "Médico", False, PRESTADOR.Text)

      Else 'Se encontrou um prestador com o dado digitado

        qBuscaPrestador.Clear
        qBuscaPrestador.Add("SELECT HANDLE")
        qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
        qBuscaPrestador.Add(" WHERE HANDLE IN (SELECT PE.PRESTADOR ")
        qBuscaPrestador.Add("                    FROM CLI_CLINICA C, ")
        qBuscaPrestador.Add("                         SAM_PRESTADOR_PRESTADORDAENTID PE ")
        qBuscaPrestador.Add("                   WHERE C.PRESTADOR = PE.ENTIDADE ")
        qBuscaPrestador.Add("                     AND C.HANDLE = :HCLINICA ")
        qBuscaPrestador.Add("       )")
        qBuscaPrestador.Add("   AND PRESTADOR LIKE '" + PRESTADOR.Text + "%'")
        qBuscaPrestador.ParamByName("HCLINICA").AsInteger = RecordHandleOfTable("CLI_CLINICA")
        qBuscaPrestador.Active = True

        vHandle = qBuscaPrestador.FieldByName("HANDLE").AsInteger
      End If
    End If
  Else 'Se não foi digitado nada, apenas abre a interface de procura

    vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
    vCriterio = "HANDLE IN (SELECT PE.PRESTADOR FROM CLI_CLINICA C, SAM_PRESTADOR_PRESTADORDAENTID PE WHERE C.PRESTADOR = PE.ENTIDADE AND C.HANDLE = " + CurrentQuery.FieldByName("CLINICA").AsString + ")"
    vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
    vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Médico", False, PRESTADOR.Text)

  End If

  Set qBuscaPrestador = Nothing
  Set Interface = Nothing

  If vHandle <>0 Then
    'CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").AsInteger = vHandle
  End If

End Sub

Public Sub RECEBEDOR_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Dim vOrdemBusca As Integer
  Dim vbAux As Boolean


  vbAux = True
  On Error GoTo caracteres
  CDbl(PRESTADOR.Text)
  vOrdemBusca = 1
  vbAux = False
caracteres:
  If vbAux Then
    vOrdemBusca = 2
  End If

  ShowPopup = False

  Dim qBuscaPrestador As Object
  Dim Interface As Object
  Set qBuscaPrestador = NewQuery
  Set Interface = CreateBennerObject("Procura.Procurar")

  If RECEBEDOR.Text <> "" Then 'Se nào foi digitado nada

    If vOrdemBusca = 2 Then 'o que foi digitado é o nome do prestador
      'Este select foi montado de acordo com o que já era executado na interface de procura
      qBuscaPrestador.Clear
      qBuscaPrestador.Add("SELECT COUNT(1) QTDE")
      qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
      qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
      qBuscaPrestador.Add("   AND Z_NOME LIKE '" + RECEBEDOR.Text + "%'")
      qBuscaPrestador.Active = True

      If (qBuscaPrestador.FieldByName("QTDE").AsInteger > 1) Or (qBuscaPrestador.FieldByName("QTDE").AsInteger = 0) Then 'A busca retornou mais de um registro, ou não retornou nenhum

        vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
        vCriterio = "RECEBEDOR = 'S'"
        vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
        vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Prestador", False, RECEBEDOR.Text)

      Else 'Se encontrou um prestador com o dado digitado

        qBuscaPrestador.Clear
        qBuscaPrestador.Add("SELECT HANDLE")
        qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
        qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
        qBuscaPrestador.Add("   AND Z_NOME LIKE '" + RECEBEDOR.Text + "%'")
        qBuscaPrestador.Active = True

        vHandle = qBuscaPrestador.FieldByName("HANDLE").AsInteger

      End If
    Else 'Foi digitado o código do prestador
      qBuscaPrestador.Clear
      qBuscaPrestador.Add("SELECT COUNT(1) QTDE")
      qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
      qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
      qBuscaPrestador.Add("   AND PRESTADOR LIKE '" + RECEBEDOR.Text + "%'")
      qBuscaPrestador.Active = True

      If (qBuscaPrestador.FieldByName("QTDE").AsInteger > 1) Or (qBuscaPrestador.FieldByName("QTDE").AsInteger = 0) Then

        vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
        vCriterio = "RECEBEDOR = 'S'"
        vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
        vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, vCampos, vCriterio, "Prestador", False, RECEBEDOR.Text)

      Else 'Se encontrou um prestador com o dado digitado

        qBuscaPrestador.Clear
        qBuscaPrestador.Add("SELECT HANDLE")
        qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
        qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
        qBuscaPrestador.Add("   AND PRESTADOR LIKE '" + RECEBEDOR.Text + "%'")
        qBuscaPrestador.Active = True

        vHandle = qBuscaPrestador.FieldByName("HANDLE").AsInteger
      End If
    End If
  Else 'Se não foi digitado nada, apenas abre a interface de procura

    vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
    vCriterio = "RECEBEDOR = 'S'"
    vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
    vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Prestador", False, RECEBEDOR.Text)

  End If

  Set Interface = Nothing
  Set qBuscaPrestador = Nothing

  '  Set Interface =CreateBennerObject("Procura.Procurar")

  '  vColunas ="PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
  '  vCriterio ="RECEBEDOR = 'S'"
  '  vCampos ="Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
  '  vHandle =Interface.Exec(CurrentSystem,"SAM_PRESTADOR",vColunas,2,vCampos,vCriterio,"Prestador",False,"")

  If vHandle <>0 Then
    'CurrentQuery.Edit
    CurrentQuery.FieldByName("RECEBEDOR").Value = vHandle
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >CurrentQuery.FieldByName("DATAFINAL").AsDateTime)And(Not CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
  bsShowMessage("A data inicial não pode ser posterior a data final!", "E")
  CanContinue = False
  DATAINICIAL.SetFocus
  Exit Sub
End If

If(CurrentQuery.FieldByName("INTERVALOAGENDA").IsNull) Or (CurrentQuery.FieldByName("INTERVALOAGENDA").AsInteger <= 0) Then
bsShowMessage("Deve ser definido um valor superior a zero para o intervalo da agenda!", "E")
CanContinue = False
Exit Sub
End If

'**************************** VERIFICA SE O RECURSO JÁ FOI CADASTRADO PARA A CLÍNICA E SE AS VIGÊNCIAS ESTÃO CRUZADAS *********************
Dim QRecurso As Object
Set QRecurso = NewQuery

QRecurso.Active = False
QRecurso.Clear
QRecurso.Add("SELECT 1 ")
QRecurso.Add("  FROM CLI_RECURSO ")
QRecurso.Add(" WHERE PRESTADOR = :PRESTADOR ")
QRecurso.Add("   AND CLINICA = :CLINICA ")
QRecurso.Add("   AND HANDLE <> :HANDLE ")

If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
  QRecurso.Add("   AND (   (DATAINICIAL <= :DATAINI AND DATAFINAL >= :DATAFIM)")
  QRecurso.Add("        OR (DATAINICIAL >  :DATAINI AND ((DATAINICIAL <= :DATAFIM AND DATAFINAL >= :DATAFIM)))")
  QRecurso.Add("        OR (DATAFINAL <  :DATAFIM AND ((DATAINICIAL <= :DATAINI AND DATAFINAL >= :DATAINI)))")
  QRecurso.Add("        OR (DATAINICIAL > :DATAINI AND DATAFINAL < :DATAFIM)")
  QRecurso.Add("        OR (DATAINICIAL <= :DATAFIM AND DATAFINAL IS NULL))")
  QRecurso.ParamByName("DATAFIM").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
Else
  QRecurso.Add("   AND (   (DATAFINAL IS NULL)")
  QRecurso.Add("        OR (DATAFINAL >= :DATAINI))")
End If


QRecurso.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
QRecurso.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
QRecurso.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
QRecurso.ParamByName("DATAINI").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
QRecurso.Active = True

If (Not QRecurso.EOF) Then
  bsShowMessage("Para médicos já cadastrados na clínica, as vigências não podem ser cruzadas!", "E")
  DATAINICIAL.SetFocus
  CanContinue = False
  Set QRecurso = Nothing
  Exit Sub
End If

Set QRecurso = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOAGENDA") Then
		BOTAOAGENDA_OnClick
	End If
	If (CommandID = "BOTAOIMPRIMIR") Then
		BOTAOIMPRIMIR_OnClick
	End If
	If (CommandID = "BOTAOSAIDA") Then
		BOTAOSAIDA_OnClick
	End If
End Sub
