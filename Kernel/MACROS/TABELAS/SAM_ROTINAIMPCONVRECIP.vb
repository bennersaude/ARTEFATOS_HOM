'HASH: 19F553830569FB484333BC7FA3CE27EF
'#Uses "*bsShowMessage"
'Macro SAM_ROTINAIMPCONVRECIP

Public Sub BOTAOCANCELA_OnClick()

  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If (Not CurrentQuery.FieldByName("USUARIOPROCESSO").IsNull) Then
    bsShowMessage("Rotina já processada.","I")
    Exit Sub
  End If

  If VisibleMode Then
  	Set Obj = CreateBennerObject("BSINTERFACE0061.Cancelar")
  	Obj.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set INTERFACE = Nothing
  Else
  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "BsBen016", _
                                    "Cancelar", _
                                    "Importação convênios de reciprocidade - Cancelar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_ROTINAIMPCONVRECIP", _
                                    "SITUACAOPROCESSO", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vmensagemretorno, "I")
  	 End If

  End If

  If VisibleMode Then
  	SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger,True,False)
  End If


  TABLE_AfterScroll
End Sub

Public Sub BOTAOIMPORTA_OnClick()

  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If (Not CurrentQuery.FieldByName("USUARIOIMPORTACAO").IsNull) Then
    bsShowMessage("Rotina já importada.","I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição","I")
    Exit Sub
  End If

  If VisibleMode Then
  	Set Obj = CreateBennerObject("BSINTERFACE0061.Importar")
  	Obj.Importar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set INTERFACE = Nothing
  Else
  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "BsBen016", _
                                    "Importar", _
                                    "Importação convênios de reciprocidade - Importar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_ROTINAIMPCONVRECIP", _
                                    "SITUACAOPROCESSO", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If

  End If

  If VisibleMode Then
  	SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger,True,False)
  End If

End Sub

Public Sub BOTAOPROCESSA_OnClick()
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If (Not CurrentQuery.FieldByName("USUARIOPROCESSO").IsNull) Then
    bsShowMessage("Rotina já processada.","I")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("USUARIOIMPORTACAO").IsNull) Then
    bsShowMessage("Rotina ainda não foi importada.","I")
    Exit Sub
  End If

   If VisibleMode Then
  	Set Obj = CreateBennerObject("BSINTERFACE0061.Processar")
  	Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set INTERFACE = Nothing
  Else
  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "BsBen016", _
                                    "Processar", _
                                    "Importação convênios de reciprocidade - Processar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_ROTINAIMPCONVRECIP", _
                                    "SITUACAOPROCESSO", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If

  End If

  If VisibleMode Then
  	SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger,True,False)
  End If

End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("DATAHORAINCLUSAO").AsDateTime = ServerNow
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("USUARIOIMPORTACAO").IsNull Then
    ARQUIVOIMPORTACAO.ReadOnly = False
  Else
    ARQUIVOIMPORTACAO.ReadOnly = True
  End If

End Sub
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELA"
			BOTAOCANCELA_OnClick
		Case "BOTAOPROCESSA"
			BOTAOPROCESSA_OnClick
		Case "BOTAOIMPORTA"
			BOTAOIMPORTA_OnClick
	End Select

End Sub
