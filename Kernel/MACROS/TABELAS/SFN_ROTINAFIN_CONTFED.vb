'HASH: 91D6878A20A4F5EBFC8B0B98DF61B2E5
'#Uses "*bsShowMessage"
'#Uses "*ProcuraPrestador"


Public Sub BOTAOCANCELAR_OnClick()
  Dim vMensagem As String
  Dim dll       As Object

  If CurrentQuery.State <>1 Then
    If VisibleMode Then
      bsShowMessage("Os parâmetros não podem estar em edição", "I")
    End If
    Exit Sub
  End If

  If VisibleMode Then
      Set dll = CreateBennerObject("BSINTERFACE0029.RotinaContribFederais")
      dll.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vMensagem)
  Else
    Dim viRetorno As Long
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Clear
    qSQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
    qSQL.Add("       CFIN.COMPETENCIA,")
    qSQL.Add("       RFIN.SEQUENCIA")
    qSQL.Add("FROM SFN_ROTINAFIN       RFIN")
    qSQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
    qSQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
    qSQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    qSQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    qSQL.Active = True

    Set dll = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = dll.ExecucaoImediata(CurrentSystem, _
                                     "SFNRECOLHIMENTO", _
                                     "RotinaContribFederais_Cancelar", _
                                     "Rotina de Recolhimento de Contribuições Federais (Cancelamento) -" + _
                                       " Competência: " + Str(Format(qSQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                       " Sequência: "   + qSQL.FieldByName("SEQUENCIA").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFIN_CONTFED", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "C", _
                                     False, _
                                     vMensagem, _
                                     Null)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
    End If

    Set qSQL = Nothing
  End If

  Set dll = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim vMensagem As String
  Dim dll       As Object

  If CurrentQuery.State <>1 Then
    If VisibleMode Then
      bsShowMessage("Os parâmetros não podem estar em edição", "I")
      Exit Sub
    End If
  End If

  If VisibleMode Then
    Set dll = CreateBennerObject("BSINTERFACE0029.RotinaContribFederais")
    dll.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vMensagem)
  Else
    Dim viRetorno As Long
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Clear
    qSQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
    qSQL.Add("       CFIN.COMPETENCIA,")
    qSQL.Add("       RFIN.SEQUENCIA")
    qSQL.Add("FROM SFN_ROTINAFIN       RFIN")
    qSQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
    qSQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
    qSQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    qSQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    qSQL.Active = True


    Set dll = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = dll.ExecucaoImediata(CurrentSystem, _
                                     "SFNRECOLHIMENTO", _
                                     "RotinaContribFederais_Processar", _
                                     "Rotina de Recolhimento de Contribuições Federais (Processamento) - " + _
                                     " Competência: " + Str(Format(qSQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                     " Sequência: "   + qSQL.FieldByName("SEQUENCIA").AsString, _
                                      CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFIN_CONTFED", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vMensagem, _
                                     Null)



      If viRetorno = 0 Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If

      Set qSQL = Nothing
   End If

   Set dll = Nothing
End Sub

Public Sub DATAAGENDA_OnPopup(ShowPopup As Boolean)
  Dim ProcurarDLL As Object
  Dim viHandle As Long
  Dim vsCampos As String
  Dim vsColunas As String
  Dim vsCriterio As String
  Dim vsTextoInicial As String

  Set ProcurarDLL = CreateBennerObject("Procura.Procurar")

  vsColunas = "SFN_IRRF_AGENDA.DATAAGENDA|SFN_IRRF_AGENDA.DATARECOLHIMENTO|SFN_IRRF_AGENDA.DATAINICIAL|SFN_IRRF_AGENDA.DATAFINAL"
  vsCampos = "Data de Agendamento|Data de Recolhimento|Data Inicial|Data Final"
  vsCriterio = "SFN_IRRF_AGENDA.TIPO = '2'"
  vsTextoInicial = DATAAGENDA.Text
  viHandle = ProcurarDLL.Exec(CurrentSystem, "SFN_IRRF_AGENDA", vsColunas, 1, vsCampos, vsCriterio, "Calendário", True, vsTextoInicial)

  If viHandle <>0 Then
    CurrentQuery.FieldByName("DATAAGENDA").AsInteger = viHandle
  End If

  Set ProcurarDLL = Nothing
  ShowPopup = False
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim viHandlePrestador As Integer
  Dim vsCPFNome As String

  If (IsNumeric(PRESTADOR.Text)) Then
    vsCPFNome = "C"
  Else
    vsCPFNome = "N"
  End If

  viHandlePrestador = ProcuraPrestador(vsCPFNome, "T", PRESTADOR.Text)

  If viHandlePrestador <> 0 Then
       CurrentQuery.Edit
       CurrentQuery.FieldByName("PRESTADOR").Value = viHandlePrestador
  End If
End Sub

Public Sub TABLE_AfterScroll()
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Início
  Dim qRotinaFin As Object
  Set qRotinaFin = NewQuery
  With qRotinaFin
    .Active = False
    .Clear
    .Add("SELECT SITUACAO")
    .Add("  FROM SFN_ROTINAFIN")
    .Add(" WHERE HANDLE = :HANDLE")
    .ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    .Active = True
  End With
  If (qRotinaFin.FieldByName("SITUACAO").AsString = "P") Or (qRotinaFin.FieldByName("SITUACAO").AsString = "S") Then
    BOTAOPROCESSAR.Enabled = False
  Else
    BOTAOPROCESSAR.Enabled = True
  End If
  Set qRotinaFin = Nothing
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Fim

  If WebMode Then
  	DATAAGENDA.WebLocalWhere = "A.TIPO = '2'"
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    bsShowMessage("Não foi possível excluir, a rotina está processada !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    bsShowMessage("Alteração negada, a rotina não está aberta !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLDATA As Object
  Dim vDATAINICIAL As Date
  Dim vDATAFINAL As Date
  Dim vDataAgenda As Date
  Dim vDATARECOLHIMENTO As Date

  Dim DI As String
  Dim DF As String
  Dim DA As String
  Dim DR As String
  Dim HANDLE As Long


  Set SQLDATA = NewQuery
  SQLDATA.Clear
  SQLDATA.Add("SELECT HANDLE,DATAINICIAL,DATAFINAL,DATAAGENDA,DATARECOLHIMENTO, TIPO FROM SFN_IRRF_AGENDA WHERE HANDLE = :HANDLE")
  SQLDATA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("DATAAGENDA").Value
  SQLDATA.Active = True
  DI = SQLDATA.FieldByName("DATAINICIAL").AsString
  DF = SQLDATA.FieldByName("DATAFINAL").AsString
  DA = SQLDATA.FieldByName("DATAAGENDA").AsString
  DR = SQLDATA.FieldByName("DATARECOLHIMENTO").AsString
  HANDLE = SQLDATA.FieldByName("HANDLE").AsString

  If HANDLE > O Then
    If SQLDATA.FieldByName("TIPO").AsString = "1" Then
      bsShowMessage("Tipo de calendário incorreto, escolha um calendário de Recolhimento de Contribuições Federais.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  vDATAINICIAL = SQLDATA.FieldByName("DATAINICIAL").Value
  vDATAFINAL = SQLDATA.FieldByName("DATAFINAL").Value
  vDataAgenda = SQLDATA.FieldByName("DATAAGENDA").Value
  vDATARECOLHIMENTO = SQLDATA.FieldByName("DATARECOLHIMENTO").Value

  If vDATAINICIAL >vDATAFINAL Then
    bsShowMessage("Data de Agendamento incorreta, a Data Final é menor que a Data Inicial.", "E")
    CanContinue = False
    Exit Sub
  End If

  If vDataAgenda >vDATARECOLHIMENTO Then
    bsShowMessage("Data de Agendamento incorreta, a Data de Recolhimento é menor que a Data de Agendamento.", "E")
    CanContinue = False
    Exit Sub
  End If


  If vDATAFINAL >vDataAgenda Then
    bsShowMessage("Data de Agendamento incorreta, a Data de Agendamento é menor que a Data Final.", "E")
    CanContinue = False
    Exit Sub
  End If

  SQLDATA.Active = False
  Set SQLDATA = Nothing
End Sub


Public Sub TABLE_NewRecord()
  Dim sql As Object
  Set sql = NewQuery

  sql.Active = False
  sql.Clear
  sql.Add("SELECT MAX(HANDLE)HANDLE ")
  sql.Add("  FROM SFN_IRRF_AGENDA   ")
  sql.Add(" WHERE TIPO = '2'        ")
  sql.Active = True

  CurrentQuery.FieldByName("DATAAGENDA").Value = sql.FieldByName("HANDLE").AsInteger

  Set sql = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
