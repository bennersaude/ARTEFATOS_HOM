'HASH: 54728F70D8A6E44615FF8D072BF11709
'Macro: SFN_ROTINARESUMO
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELARREC_OnClick()

  If VisibleMode Then
    	Dim interface As Object

    	Set interface = CreateBennerObject("BSInterface0021.CancelarRec")

	    interface.Exec(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)
        	'WriteAudit("C", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Recebimento - Cancelamento")
    	Set interface = Nothing

  ElseIf WebMode Then

 		Dim vsMensagemErro As String
    	Dim viRetorno As Long


	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     	"BSFIN006", _
                                     	"RotinaResumo_CancelarRec", _
                                     	"Rotina de Cancelamento de resumo de IRRF de recebimento: " + CurrentQuery.FieldByName("DESCRICAO").AsString, _
                                     	CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     	"SFN_ROTINARESUMO", _
                                     	"SITUACAOPROCESSAMENTO", _
                                     	"", _
                                     	"", _
                                     	"C", _
                                     	True, _
	                                   	vsMensagemErro, _
                                     	Null)

	    If viRetorno = 0 Then
      		bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If
  End If
End Sub

Public Sub BOTAODIRF_OnClick()
  If VisibleMode Then

  	Dim interface As Object

  	Set interface = CreateBennerObject("BSINTERFACE0021.ExportaDIRF")
  	interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  	Set interface = Nothing

  '	Select Case CurrentQuery.FieldByName("TABTIPO").AsInteger
	'    Case 1
     ' 	WriteAudit("G", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Pagamento - Geração de arquivo DIRF")
    '	Case 2
     ' 	WriteAudit("G", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Recebimento - Geração de arquivo DIRF")
  	'End Select

  	RefreshNodesWithTable("SFN_ROTINARESUMO")

  End If

End Sub

Public Sub BOTAOEXPORTAR_OnClick()

  If VisibleMode Then

  	Dim interface As Object

  	Set interface = CreateBennerObject("BSINTERFACE0021.ExportacaoPag")
  	interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  	Set interface = Nothing

  	'WriteAudit("E", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Pagamento - Exportação")

  	RefreshNodesWithTable("SFN_ROTINARESUMO")

  End If

End Sub

Public Sub BOTAOEXPORTARREC_OnClick()
  If VisibleMode Then
  	Dim interface As Object

  	Set interface = CreateBennerObject("BSINTERFACE0021.ExportacaoRec")
  	interface.Exec(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)
  	Set interface = Nothing


   	'WriteAudit("E", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Recebimento - Exportação")

  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("BSINTERFACE0021.ProcessarPag")

  interface.Exec(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)
'  WriteAudit("P", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Pagamento - Processamento")

  Set interface = Nothing
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.FieldByName("CANCELADODATA").IsNull Then
  	If VisibleMode Then

    	Dim interface As Object
    	Set interface = CreateBennerObject("BSINTERFACE0021.CancelarPag")
	    interface.Exec(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)
        	'WriteAudit("C", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Pagamento - Cancelamento")
    	Set interface = Nothing

    ElseIf WebMode Then

 		Dim vsMensagemErro As String
    	Dim viRetorno As Long


	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     	"BSFIN006", _
                                     	"RotinaResumo_CancelarPag", _
                                     	"Rotina de Cancelamento de resumo de IRRF: " + CurrentQuery.FieldByName("DESCRICAO").AsString, _
                                     	CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     	"SFN_ROTINARESUMO", _
                                     	"SITUACAOPROCESSAMENTO", _
                                     	"", _
                                     	"", _
                                     	"C", _
                                     	True, _
	                                   	vsMensagemErro, _
                                     	Null)

	    If viRetorno = 0 Then
      		bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If

    End If
  Else
    bsShowMessage("A rotina já foi cancelada", "I")
    Exit Sub
  End If

End Sub

Public Sub BOTAOPROCESSARREC_OnClick()
  Dim interface As Object
  	Set interface = CreateBennerObject("BSINTERFACE0021.ProcessarRec")

	interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,CurrentQuery.FieldByName("GRUPOCONTRATO").AsString,CurrentQuery.FieldByName("CONTRATO").AsString)
      'WriteAudit("P", HandleOfTable("SFN_ROTINARESUMO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Resumo IRRF Recebimento - Processamento")
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	If CurrentQuery.FieldByName("TABTIPO").AsInteger = 1 Then
		BOTAOEXPORTAR.Visible = True
		BOTAOEXPORTARREC.Visible = False

		BOTAOPROCESSAR.Visible = True
		BOTAOPROCESSARREC.Visible = False

		BOTAOCANCELAR.Visible = True
		BOTAOCANCELARREC.Visible = False
	Else
		BOTAOEXPORTAR.Visible = False
		BOTAOEXPORTARREC.Visible = True

		BOTAOPROCESSAR.Visible = False
		BOTAOPROCESSARREC.Visible = True

		BOTAOCANCELAR.Visible = False
		BOTAOCANCELARREC.Visible = True
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("PROCESSADODATA").IsNull Then
    CanContinue = False
    bsShowMessage("A rotina está processada!", "E")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT CONTROLADOTORC FROM SFN_PARAMETROSFIN")
  SQL.Active = True

  If (SQL.FieldByName("CONTROLADOTORC").AsString = "2") Then
    If (CurrentQuery.FieldByName("TESOURARIA").IsNull) Then
       bsShowMessage("Não é possível criar rotinas de recebimento sem tesouraria, se o sistema estiver parametrizado para utilizar dotação orçamentária!", "E")
       CanContinue = False
       Exit Sub
    End If
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_NewRecord()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT NOME, RAZAO FROM EMPRESAS WHERE HANDLE = :EMPRESACORRENTE")
  SQL.ParamByName("EMPRESACORRENTE").AsInteger = CurrentCompany
  SQL.Active = True

  CurrentQuery.FieldByName("NOMEEMPRESA").AsString = SQL.FieldByName("RAZAO").AsString

  Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	Select Case CommandID

		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick

		Case "BOTAOCANCELARREC"
			BOTAOCANCELARREC_OnClick

	End Select

End Sub
