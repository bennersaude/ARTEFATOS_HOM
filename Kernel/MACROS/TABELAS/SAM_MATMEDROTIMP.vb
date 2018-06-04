'HASH: 44D579DB1C46ABEE3316EEEC21684B6E
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()

	Dim interface As Object

	If CurrentQuery.State <>1 Then
		bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
		Exit Sub
	End If

    If Not ValidarDataRotina Then
      bsShowMessage("Não pode ser cancelada a rotina pois existe outra rotina criada com data superior ou igual a esta rotina favor verificar!", "I")
      Exit Sub
    End If


    If (WebMode) Then
        Dim vsMensagemErro As String
        Dim viRetorno As Long

        Set interface = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = interface.ExecucaoImediata(CurrentSystem, _
                                     "CliImpMedicamento", _
                                     "Cancelar", _
                                     "Rotina de Importação de Materiais e Medicamentos (Cancelamento)", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_MATMEDROTIMP", _
                                     "SITUACAOIMP", _
                                     "", _
                                     "", _
                                     "C", _
                                     False, _
                                     vsMensagemErro, _
                                     Null)

        If viRetorno = 0 Then
          bsShowMessage("Processo enviado para execução no servidor!", "I")
        Else
          bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
        End If
        Set interface = Nothing
    Else
      Set interface = CreateBennerObject("BSINTERFACE0045.IMPORTAR")
      interface.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    End If
    RefreshNodesWithTable("CLI_ROTIMPARQUIVO")
    Set interface = Nothing

	'Luciano T. Alberti - SMS 108011 - 02/03/2009 - Início
	If VisibleMode Then
		RefreshNodesWithTable("SAM_MATMEDROTIMP")
	End If
	'Luciano T. Alberti - SMS 108011 - 02/03/2009 - Fim
End Sub


Public Sub BOTAOPROCESSAR_OnClick()
	Dim interface      As Object
	Dim vSequencia     As String
	Dim qAchaSequencia As Object
	Dim vsMensagemErro As String
	Dim viRetorno      As Long

    Dim sx As CSServerExec
    Dim SQL As Object
	Set sx = NewServerExec
	Set SQL = NewQuery

	If CurrentQuery.State <> 1 Then
		bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "E")
		Exit Sub
	End If

    If Not ValidarDataRotina Then
      bsShowMessage("Não pode ser processada a rotina pois existe outra rotina criada com data superior ou igual a esta rotina favor verificar!", "I")
      Exit Sub
    End If


	If (Not ValidarOrigem) Then
      Exit Sub
	End If

    If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 4 Then
  		sx.Description = "Rotina de Importação de Materias e Medicamentos - SIMPRO (XML)"
  		sx.DllClassName = "Benner.Saude.Services.ProcContas.ImportarMatMedSimproXML"
  		sx.SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
  		sx.Execute

        SQL.Clear
        SQL.Add("UPDATE SAM_MATMEDROTIMP SET SITUACAOIMP = '2' WHERE HANDLE = :HANDLE")
        SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQL.ExecSQL

	    bsShowMessage("Processo enviado para execução no servidor!", "I")
	Else
		If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 3 Then
  			'Set interface = CreateBennerObject("Benner.Saude.Services.ProcContas.ImportarMatMedBrasindice")
    	    'SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
		    'interface.Exec(CurrentSystem)

  			sx.Description = "Rotina de Importação de Materias e Medicamentos - BRASINDICE"
	  		sx.DllClassName = "Benner.Saude.Services.ProcContas.ImportarMatMedBrasindice"
  			sx.SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
	  		sx.Execute

    	    SQL.Clear
	        SQL.Add("UPDATE SAM_MATMEDROTIMP SET SITUACAOIMP = '2' WHERE HANDLE = :HANDLE")
    	    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	        SQL.ExecSQL

		    bsShowMessage("Processo enviado para execução no servidor!", "I")
	    End If
	End If


    RefreshNodesWithTable("CLI_ROTIMPARQUIVO")
    Set interface = Nothing
	Set sx = Nothing

	'Luciano T. Alberti - SMS 108011 - 02/03/2009 - Início
	If VisibleMode Then
		RefreshNodesWithTable("SAM_MATMEDROTIMP")
	End If
	'Luciano T. Alberti - SMS 108011 - 02/03/2009 - Fim

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  If CurrentQuery.FieldByName("SITUACAOIMP").Value = "5" Then
    bsShowMessage("Impossível excluir registro porque rotina já foi processada!", "E")
    CanContinue = False
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If CurrentQuery.FieldByName("SITUACAOIMP").Value = "5" Then
    bsShowMessage("Impossível alterar campo porque rotina já foi processada!", "E")
    CanContinue = False
  End If

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vEdicao As String
  Dim QueryBuscaEdicao As Object
  Dim qverifica As Object

  If CurrentQuery.FieldByName("TABORIGEM").AsString = "4" Then
    If CurrentQuery.FieldByName("ARQSIMPRO").IsNull Then
      bsShowMessage("Arquivo SIMPRO deve ser selecionado!", "E")
      CanContinue = False
      Exit Sub
    Else
	  If UCase(Right(CurrentQuery.FieldByName("ARQSIMPRO").AsString, 4)) <> ".XML" Then
        bsShowMessage("Arquivo SIMPRO deve ter extensão .XML!", "E")
        CanContinue = False
        Exit Sub
	  End If
	End If
  End If

  If CurrentQuery.State = 3 Or CurrentQuery.State = 2 Then 'Inserir ou Editar
    If Not ValidarDataRotina Then
      bsShowMessage("Não pode ser salva a rotina pois existe outra rotina criada com data superior ou igual a esta rotina favor verificar!", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 4 Then
    CurrentQuery.FieldByName("ARQSIMPRO").AsString  = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(TiraAcento(CurrentQuery.FieldByName("ARQSIMPRO").AsString,  True), " ", ""), "*",""),"!",""),"@",""),"#",""), "%",""), "&",""), "(",""), ")",""), "-","")
  End If

  While Len(CurrentQuery.FieldByName("EDICAO").AsString) < 5
    CurrentQuery.FieldByName("EDICAO").AsString = "0" + CurrentQuery.FieldByName("EDICAO").AsString
  Wend

  CanContinue = ValidarOrigem

  If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 3 Then
    If CurrentQuery.FieldByName("ARQUIVOMATERIAL").IsNull _
    	And CurrentQuery.FieldByName("ARQUIVOMEDICAMENTO").IsNull _
    	And CurrentQuery.FieldByName("ARQUIVOSOLUCAO").IsNull _
    	And CurrentQuery.FieldByName("ARQUIVOMEDICAMENTORESTRITOHOSP").IsNull Then
      bsShowMessage("Ao menos um arquivo deve ser selecionado!", "E")
      CanContinue = False
    End If

    If CurrentQuery.FieldByName("ESTADO").AsString <> "ZF" Then
      Dim sql As Object
      Set sql = NewQuery
      sql.Clear
      sql.Add("SELECT COUNT(1) QTD FROM ESTADOS WHERE SIGLA = :SIGLA")
      sql.ParamByName("SIGLA").AsString = CurrentQuery.FieldByName("ESTADO").AsString
      sql.Active = True

      If sql.FieldByName("QTD").AsInteger = 0 Then
        bsShowMessage("A sigla do estado é inválida!", "E")
        CanContinue = False
      End If
      Set sql = Nothing
    End If

  End If

  Set qverifica = NewQuery
  If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 3 Then
    qverifica.Clear
    qverifica.Add("SELECT FECHAVIGENCIA FROM SAM_MATMED_TABGENESTADO WHERE ESTADO = :ESTADO")
    qverifica.ParamByName("ESTADO").AsString = CurrentQuery.FieldByName("ESTADO").AsString
    qverifica.Active = True

    If qverifica.FieldByName("FECHAVIGENCIA").AsString = "S" And _
       (CurrentQuery.FieldByName("ARQUIVOMATERIAL").IsNull Or _
    	CurrentQuery.FieldByName("ARQUIVOMEDICAMENTO").IsNull Or _
    	CurrentQuery.FieldByName("ARQUIVOSOLUCAO").IsNull Or _
    	CurrentQuery.FieldByName("ARQUIVOMEDICAMENTORESTRITOHOSP").IsNull) Then

	  bsShowMessage("Necessário informar os quatro arquivos quando o parâmetro 'Fechar vigência do registro' estiver marcado para o estado informado", "E")
      CanContinue = False
      Exit Sub

    End If
  End If
  Set qverifica = Nothing

End Sub

Public Sub TABLE_NewRecord()
  'Dim qAchaSequencia As Object
  'Set qAchaSequencia = NewQuery
  'qAchaSequencia.Active = False
  'qAchaSequencia.Add("SELECT COUNT(HANDLE) AS SEQUENCIA ")
  'qAchaSequencia.Add("  FROM SAM_MATMEDROTIMP ")
  'qAchaSequencia.Add(" WHERE TABORIGEM=2")
  'qAchaSequencia.Active = True
  'If qAchaSequencia.FieldByName("SEQUENCIA").AsInteger = 0 Then
  '  vSequencia = "00001"
  'Else
  '  vSequencia = Format(Str(qAchaSequencia.FieldByName("Sequencia").AsInteger + 1), "00000")
  'End If
  'Set qAchaSequencia = Nothing
  'CurrentQuery.FieldByName("EDICAO").AsString = vSequencia

  CurrentQuery.FieldByName("EDICAO").AsString = "00000"
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

 If CommandID = "BOTAOPROCESSAR" Then
    BOTAOPROCESSAR_OnClick
 ElseIf CommandID = "BOTAOCANCELAR" Then
    BOTAOCANCELAR_OnClick
 End If

End Sub
Public Function ValidarOrigem As Boolean

  ValidarOrigem = True

  If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 1 Then
      bsShowMessage("Os arquivos de origem 'Brasíndice 705 ou inferior' não são mais suportados. Para importar arquivos do brasíndice, utilizar a origem 'Brasíndice'.", "E")
	  ValidarOrigem = False
  Else
    If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 2 Then
      bsShowMessage("Os arquivos de origem 'Simpro (Antigo)' não são mais suportados. Para importar arquivos do simpro, utilizar a origem 'Simpro (XML)'.", "E")
	  ValidarOrigem = False
	End If
  End If

End Function

Public Function ValidarDataRotina As Boolean
  ValidarDataRotina = True
  Dim qverifica As Object
  Set qverifica = NewQuery
  If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 3 Then
    qverifica.Clear
    qverifica.Add("SELECT COUNT(1) QTD              ")
    qverifica.Add("  FROM SAM_MATMEDROTIMP          ")
    qverifica.Add(" WHERE DATAROTINA >= :DATAROTINA ")
    qverifica.Add("   AND TABORIGEM = :TABORIGEM    ")
    qverifica.Add("   AND ESTADO = :ESTADO          ")
    qverifica.Add("   AND HANDLE <> :HANDLE         ")
    qverifica.ParamByName("DATAROTINA").AsDateTime = CurrentQuery.FieldByName("DATAROTINA").AsDateTime
    qverifica.ParamByName("TABORIGEM").AsInteger = CurrentQuery.FieldByName("TABORIGEM").AsInteger
    qverifica.ParamByName("ESTADO").AsString = CurrentQuery.FieldByName("ESTADO").AsString
    qverifica.ParamByName("HANDLE").AsString = CurrentQuery.FieldByName("HANDLE").AsInteger
    qverifica.Active = True
    If qverifica.FieldByName("QTD").AsInteger > 0 Then
      ValidarDataRotina = False
    End If
  End If

  If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 4 Then
    qverifica.Clear
    qverifica.Add("SELECT COUNT(1) QTD                     ")
	qverifica.Add("  FROM SAM_MATMEDROTIMP                 ")
	qverifica.Add(" WHERE DATAROTINA    >= :DATAROTINA     ")
	qverifica.Add("   AND TABORIGEM      = :TABORIGEM      ")
	qverifica.Add("   AND TABELAALIQUOTA = :TABELAALIQUOTA ")
	qverifica.Add("   AND HANDLE        <> :HANDLE         ")
	qverifica.ParamByName("DATAROTINA").AsDateTime    = CurrentQuery.FieldByName("DATAROTINA").AsDateTime
	qverifica.ParamByName("TABORIGEM").AsInteger      = CurrentQuery.FieldByName("TABORIGEM").AsInteger
	qverifica.ParamByName("TABELAALIQUOTA").AsInteger = CurrentQuery.FieldByName("TABELAALIQUOTA").AsInteger
	qverifica.ParamByName("HANDLE").AsInteger         = CurrentQuery.FieldByName("HANDLE").AsInteger
    qverifica.Active = True
    If qverifica.FieldByName("QTD").AsInteger > 0 Then
      ValidarDataRotina = False
    End If
  End If
  Set qverifica = Nothing
End Function
Public Sub TABLE_AfterScroll()
  If CurrentQuery.State = 1 Then ' leitura
    If CurrentQuery.FieldByName("SITUACAOIMP").AsString = "5" Then
	  BOTAOPROCESSAR.Enabled = False
      BOTAOCANCELAR.Enabled = True
    ElseIf CurrentQuery.FieldByName("SITUACAOIMP").AsString = "1" Then
      BOTAOCANCELAR.Enabled = False
      BOTAOPROCESSAR.Enabled = True
    Else
      BOTAOPROCESSAR.Enabled = False
      BOTAOCANCELAR.Enabled = False
    End If
  Else
    BOTAOPROCESSAR.Enabled = False
    BOTAOCANCELAR.Enabled = False
  End If
End Sub
