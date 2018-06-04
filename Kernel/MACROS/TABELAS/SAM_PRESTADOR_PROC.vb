'HASH: 9DC629BDBE088AE0969F04F97BBEAEA9
'Macro: SAM_PRESTADOR_PROC

'#Uses "*bsShowMessage"

Option Explicit
Dim vHandleCreden As Long
Dim vTipoProcesso As String
Dim EstadodaTabela As Long

Public Sub BOTAOALTERARRESPONSAVEL_OnClick()
   If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
     If (Not InTransaction) Then
       StartTransaction
     End If
     Dim qUpdate As Object
     Set qUpdate = NewQuery
     qUpdate.Add("UPDATE SAM_PRESTADOR_PROC ")
     qUpdate.Add("   SET RESPONSAVEL = :RESPONSAVEL ")
     qUpdate.Add(" WHERE HANDLE = :HANDLE ")
     qUpdate.ParamByName("RESPONSAVEL").AsInteger = CurrentUser
     qUpdate.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
     qUpdate.ExecSQL
     Set qUpdate = Nothing
     If (InTransaction) Then
       Commit
     End If
     RefreshNodesWithTable("SAM_PRESTADOR_PROC")
   Else
     bsShowMessage("Processo finalizado! Operação não permitida.", "I")
   End If
End Sub

Public Sub BOTAOCARTA_OnClick()
  Dim OLESamCarta As Object
  Set OLESamCarta = CreateBennerObject("SamCarta.Impressao")
  OLESamCarta.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set OLESamCarta = Nothing
End Sub

Public Sub BOTAOCONSULTARPRESTADOR_OnClick()
  Dim vPrestador
  Dim Interface As Object
  Set Interface = CreateBennerObject("CA005.ConsultaPrestador")
  vPrestador = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  Interface.info(CurrentSystem, vPrestador)
  Set Interface = Nothing
End Sub

Public Sub BOTAOFINALIZAR_OnClick()
 Dim vsMensagemErro As String
 Dim viRetorno As Long
 Dim dllBSServerExec As Object
 Dim Interface As Object
 Dim vsMensagem As String
 Dim vcContainer As CSDContainer
 Dim qPrestadorSubstituto As BPesquisa
 Dim qTipoPrestador As BPesquisa
 Dim SamPrestadorBLL As CSBusinessComponent

 Set vcContainer = NewContainer
 Set qPrestadorSubstituto = NewQuery
 Set qTipoPrestador = NewQuery

 If CurrentQuery.State <> 1 Then
   bsShowMessage("O registro corrente deve ser gravado primeiro. ", "I")
   Exit Sub
 End If

 If (CurrentQuery.FieldByName("TIPOPROCESSO").AsString = "C") Then
   vsMensagem = "Credenciamento de Prestador"
 Else
   vsMensagem = "Descredenciamento de Prestador"
 End If

 vcContainer.AddFields("HANDLE:INTEGER;MENSAGEM:STRING")

 vcContainer.Insert
 vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 vcContainer.Field("MENSAGEM").AsString = vsMensagem

 If (CurrentQuery.FieldByName("TIPOPROCESSO").AsString = "D") Then 'Descredenciamento

   qTipoPrestador.Add("SELECT TP.OBRIGASUBSTITUICAO                              ")
   qTipoPrestador.Add("  FROM SAM_PRESTADOR P                                    ")
   qTipoPrestador.Add("  JOIN SAM_TIPOPRESTADOR TP ON P.TIPOPRESTADOR = TP.HANDLE")
   qTipoPrestador.Add(" WHERE P.HANDLE = :PRESTADOR                              ")
   qTipoPrestador.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
   qTipoPrestador.Active = True

   If (qTipoPrestador.FieldByName("OBRIGASUBSTITUICAO").AsString = "S")Then

     qPrestadorSubstituto.Add("SELECT HANDLE                   ")
	 qPrestadorSubstituto.Add("FROM SAM_PRESTADOR_SUBSTITUTO   ")
	 qPrestadorSubstituto.Add("WHERE PRESTADORPROC = :PROCESSO ")
	 qPrestadorSubstituto.ParamByName("PROCESSO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	 qPrestadorSubstituto.Active = True

     If qPrestadorSubstituto.EOF Then
       bsShowMessage("Deve haver uma Substituição para finalizar o processo de descredenciamento do prestador.", "I")
       Set qPrestadorSubstituto = Nothing
       Set qTipoPrestador = Nothing
       Exit Sub

     End If

   End If

   Set qPrestadorSubstituto = Nothing
   Set qTipoPrestador = Nothing

   If (VisibleMode) Then
     Set Interface = CreateBennerObject("BSINTERFACE0032.PROCESSAR")
     Interface.DESCREDENCIAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
   Else
     Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
     viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                  "BSPRE012", _
                                                  "PROCESSO_DESCREDENCIAR", _
                                                  vsMensagem, _
                                                  0, _
                                                  "SAM_PRESTADOR_PROC", _
                                                  "SITUACAOPROCESSAMENTO", _
                                                  "", _
                                                  "", _
                                                  "P", _
                                                  False, _
                                                  vsMensagemErro, _
                                                  vcContainer)
     If viRetorno = 0 Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
     Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
     End If
  End If
 Else 'Credenciamento
   If (VisibleMode) Then
     Set Interface = CreateBennerObject("BSINTERFACE0032.PROCESSAR")
     Interface.CREDENCIAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
   Else
     Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
     viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                  "BSPRE012", _
                                                  "PROCESSO_CREDENCIAR", _
                                                  vsMensagem, _
                                                  0, _
                                                  "SAM_PRESTADOR_PROC", _
                                                  "SITUACAOPROCESSAMENTO", _
                                                  "", _
                                                  "", _
                                                  "P", _
                                                  False, _
                                                  vsMensagemErro, _
                                                  vcContainer)

     If viRetorno = 0 Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
     Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
   End If
 End If

End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser
End Sub

Public Sub TABLE_AfterPost()
  RESPONSAVEL.ReadOnly = True

  If (EstadodaTabela = 3) And (CurrentQuery.FieldByName("TIPOPROCESSO").AsString = "D") Then

    If (VisibleMode) Then
      Dim Interface As Object
      Set Interface = CreateBennerObject("SamVaga.Atendimento")
      Interface.ChkVaga(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
      Set Interface = Nothing
    Else
      Dim Param As Object
      Set Param = NewQuery

      Param.Active = False
      Param.Clear
      Param.Add("SELECT CONTABBENEFVAGAS        ")
      Param.Add("  FROM SAM_PARAMETROSPRESTADOR ")
      Param.Active = True

      If (Param.FieldByName("CONTABBENEFVAGAS").AsString = "S") Then
        Dim spProc As BStoredProc
        Set spProc = NewStoredProc
        spProc.Name = "BSPRE_CHKVAGA"
        spProc.AddParam("p_Prestador", ptInput, ftInteger)
        spProc.AddParam("p_Processo", ptInput, ftInteger)
        spProc.ParamByName("p_Prestador").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
     	spProc.ParamByName("p_Processo").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        spProc.ExecProc
        Set spProc = Nothing
      End If
    End If
  End If
End Sub

Public Sub TABLE_AfterScroll()
  DATAFINAL.ReadOnly = True
  If CurrentQuery.State = 3 Then
    RESPONSAVEL.ReadOnly = False
  Else
    If Not CurrentQuery.FieldByName("RESPONSAVEL").IsNull Then
      RESPONSAVEL.ReadOnly = True
    End If
  End If

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAINICIAL.ReadOnly = True
  Else
    DATAINICIAL.ReadOnly = False
  End If

  If (VisibleMode And NodeInternalCode = 2)  Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_579") Or WebMode Then
    'GRUPOCREDENCIAR.ReadOnly = False
    FECHARVIGENCIATABELAPRECO.ReadOnly = False

    'GRUPODESCREDENCIAMENTO.ReadOnly = True
    FECHARVIGENCIAS.ReadOnly = True
    DESCRENDENCIAFILIAIS.ReadOnly = True
    DESCREDENCIARCOMFATURASABERTAS.ReadOnly = True

    If CurrentQuery.State = 3 Then
      CurrentQuery.FieldByName("TIPOPROCESSO").AsString = "C"
    End If
  End If

  If (VisibleMode And NodeInternalCode = 3) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_578") Then
   'GRUPOCREDENCIAR.ReadOnly = False
    FECHARVIGENCIATABELAPRECO.ReadOnly = True

    'GRUPODESCREDENCIAMENTO.ReadOnly = True
    FECHARVIGENCIAS.ReadOnly = False
    DESCRENDENCIAFILIAIS.ReadOnly = False
    DESCREDENCIARCOMFATURASABERTAS.ReadOnly = False
    If CurrentQuery.State = 3 Then
      CurrentQuery.FieldByName("TIPOPROCESSO").AsString = "D"
    End If

  End If


End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  Dim VAGA As Object
  Dim DESCR As Object
  Set VAGA = NewQuery
  Set DESCR = NewQuery

  VerFimProc CanContinue

  If CanContinue Then

    VAGA.Add("DELETE FROM SAM_PRESTADOR_PROC_DESCRE_VAGA")
    VAGA.Add("WHERE PRESTADORPROCESSO In ( Select HANDLE FROM SAM_PRESTADOR_PROC_DESCRE ")
    VAGA.Add("                             WHERE PRESTADORPROCESSO In ( Select PROCESSO FROM SAM_PRESTADOR_PROC")
    VAGA.Add("                                                          WHERE HANDLE = :HANDLE))")
    VAGA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    VAGA.ExecSQL

    DESCR.Add("DELETE  FROM SAM_PRESTADOR_PROC_DESCRE")
    DESCR.Add("WHERE PRESTADORPROCESSO In ( Select PROCESSO FROM SAM_PRESTADOR_PROC")
    DESCR.Add("                             WHERE HANDLE = :HANDLE )")
    DESCR.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    DESCR.ExecSQL

  End If

  Set VAGA = Nothing
  Set DESCR = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String


  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  VerFimProc CanContinue
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT FILIALPADRAO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR")
  SQL.Active = True

  If SQL.FieldByName("FILIALPADRAO").IsNull Then
    Exit Sub
  End If

  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub




Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim vDataI As String
  Dim vDataF As String
  Dim Linha As String
  Dim vTipoProcesso As String
  Dim SQL As Object
  EstadodaTabela = CurrentQuery.State
  'VERIFICAR SE EXISTE OUTRO PROCESSO EM ABERTO
  Set SQL = NewQuery

	If (VisibleMode And NodeInternalCode = 2) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_579") Or WebMode Then
      vTipoProcesso = "C"
    End If
    If (VisibleMode And NodeInternalCode = 3) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_578") Then
      vTipoProcesso = "D"
    End If

  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC A WHERE A.HANDLE <> :HANDLE AND A.PRESTADOR = :PRESTADOR AND A.DATAFINAL IS NULL")
  SQL.Add("AND TIPOPROCESSO =:TIPOPROCESSO")
  SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("TIPOPROCESSO").AsString = vTipoProcesso
  SQL.Active = True
  If Not SQL.EOF Then
    'Cancontinue =False
    bsShowMessage("Atenção! Existe outro processo em aberto!", "I")
    'Exit Sub
  End If
  'veririficar se existe processos com data inicial ou final maior que a data inicial do processo corrente
  SQL.Clear
  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC A WHERE (A.PRESTADOR = :PRESTADOR) AND")
  SQL.Add("(HANDLE <> :HANDLE) AND")
  SQL.Add("(:DATAI < DATAINICIAL OR :DATAI < DATAFINAL)")
  SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("DATAI").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  SQL.Active = True
  If Not SQL.EOF Then
    'Cancontinue =False
    'MsgBox "Vigência inválida- verifique outros processos!"
    'Exit Sub
  End If


  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_PRESTADOR_CRED WHERE PRESTADOR = :PRESTADOR AND")
  SQL.Add(" DESCREDENCIAMENTODATA IS NULL ")
  SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.Active = True
  If SQL.EOF And vTipoProcesso = "D" Then
    CanContinue =False
    bsShowMessage("Prestador não está credenciado.", "I")
    Exit Sub
  End If


    If (VisibleMode And NodeInternalCode = 2)  Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_579") Or WebMode Then
      vTipoProcesso = "C"
    End If
    If (VisibleMode And NodeInternalCode = 3) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_578") Then
      vTipoProcesso = "D"
    End If

  CurrentQuery.FieldByName("TIPOPROCESSO").Value = vTipoProcesso

  'sms 108068 crislei.sorrilha faz a validação dos campos obrigatorios
  If vTipoProcesso = "C" Then
	If CurrentQuery.FieldByName("FECHARVIGENCIATABELAPRECO").AsBoolean = False Then
	  bsShowMessage("Campo Fechar as Vigências das Regras de Preços é obrigatório","E")
	  CanContinue = False
      Exit Sub
	End If
  End If
  If vTipoProcesso = "D" Then
	If CurrentQuery.FieldByName("DESCREDENCIARCOMFATURASABERTAS").AsBoolean = False Then
	  bsShowMessage("Campo Descredenciar com faturas abertas é obrigatório","E")
	  CanContinue = False
      Exit Sub
	End If
  End If

  If CurrentQuery.State = 3 Then
    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime >CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
      CanContinue = False
      CurrentQuery.Cancel
      bsShowMessage("Data Final maior que data inicial", "E")
      Exit Sub
    End If

    'Mauricio
    'If CurrentQuery.FieldByName("TIPOPROCESSO").Value ="D" Then
    '//verificar se o prestador já não está descredenciado
    Dim Prest As Object
    Set Prest = NewQuery
    Prest.Add("SELECT DATADESCREDENCIAMENTO,MOTIVODESCREDENCIAMENTO FROM SAM_PRESTADOR WHERE HANDLE = :PRESTADOR")
    Prest.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
    Prest.Active = True
    If CurrentQuery.FieldByName("TIPOPROCESSO").Value = "D" Then
      If Not Prest.FieldByName("DATADESCREDENCIAMENTO").IsNull Then
        bsShowMessage("Prestador já está descredenciado - Operação incoerente!", "E")
        CanContinue = False
        Exit Sub
      End If
    End If

    If Not Prest.FieldByName("MOTIVODESCREDENCIAMENTO").IsNull Then

      Dim S As Object
      Set S = NewQuery

      S.Add("SELECT PERMITERECREDENCIAMENTO FROM SAM_MOTIVODESCREDENCIAMENTO WHERE HANDLE = :HANDLE")
      S.ParamByName("HANDLE").Value = Prest.FieldByName("MOTIVODESCREDENCIAMENTO").AsInteger
      S.Active = True

      If S.FieldByName("PERMITERECREDENCIAMENTO").Value = "N" Then
        CanContinue = False
        bsShowMessage("Motivo do descredenciamento não permite re-credenciamento", "E")
        'CurrentQuery.Cancel
        Exit Sub
      End If

    End If


  End If

  RefreshNodesWithTable("SAM_PRESTADOR_PROC")

End Sub

Public Sub VerFimProc(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage("Processo finalizado!Operação não permitida.", "E")
  End If
End Sub

Public Function TemFilhos As Boolean
  Dim SQL As Object
  Dim S As String

    If (VisibleMode And NodeInternalCode = 2)  Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_579") Then
      S = "SELECT COUNT(*) TOT FROM SAM_PRESTADOR_PROC_CREDEN WHERE PRESTADORPROCESSO = :PROC"
    End If
    If (VisibleMode And NodeInternalCode = 3) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PROC_578") Then
      S = "SELECT COUNT(*) TOT FROM SAM_PRESTADOR_PROC_DESCRE WHERE PRESTADORPROCESSO = :PROC"
    End If
  Set SQL = NewQuery
  SQL.Add(S)
  SQL.ParamByName("PROC").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  TemFilhos = IIf(SQL.FieldByName("TOT").AsInteger >0, True, False)
  Set SQL = Nothing
End Function


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

 If CommandID = "BOTAOFINALIZAR" Then
    BOTAOFINALIZAR_OnClick
 ElseIf CommandID = "BOTAOALTERARRESPONSAVEL" Then
    BOTAOALTERARRESPONSAVEL_OnClick
 End If

End Sub

Public Sub EDITAL_OnPopup(ShowPopup As Boolean)
	If Not PossuiEnderecoAtendimento Then
		bsShowMessage("Não é possível localizar um edital pois o prestador não possui endereço de atendimento ou o endereço de atendimento está cancelado.", "I")
		ShowPopup = False
		Exit Sub
	End If
	Exit Sub
End Sub

Public Function PossuiEnderecoAtendimento As Boolean
	Dim SQL As Object

	Set sql = NewQuery

	sql.Clear

	sql.Add("SELECT COUNT(1) aCHOUEND                                                   ")
	sql.Add("  FROM SAM_PRESTADOR_ENDERECO E                                            ")
	sql.Add(" WHERE E.PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString )
	sql.Add("   AND E.ATENDIMENTO   = 'S'                                               ")
	sql.Add("   AND E.DATACANCELAMENTO  IS NULL                                         ")

	sql.Active = True

	If sql.FieldByName("aCHOUEND").AsInteger = 0 Then
		PossuiEnderecoAtendimento = False
	Else
		PossuiEnderecoAtendimento = True
	End If

	Set sql = Nothing

End Function
