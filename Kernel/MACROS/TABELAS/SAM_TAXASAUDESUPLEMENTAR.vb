'HASH: 7A5C1EDB321EFA52AF52B7424B46330C
'Macro: SAM_TAXASAUDESUPLEMENTAR
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
  Dim Interface As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Integer

  If VisibleMode Then
  	Set Interface = CreateBennerObject("BSINTERFACE0057.Cancelar")
  	Interface.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,vsMensagemRetorno)
  	Set Interface = Nothing
  Else
  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "SamTaxaSaudeSuplementar", _
                                    "Cancelar", _
                                    "TSS - Taxa Saúde Suplementar - Cancelar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_TAXASAUDESUPLEMENTAR", _
                                    "SITUACAOPROCESSO", _
                                    "", _
                                    "", _
                                    "C", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)
    If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	End If

  End If


  If Not WebMode Then
  	RefreshNodesWithTable("SAM_TAXASAUDESUPLEMENTAR")
  End If
End Sub


Public Sub BOTAOPROCESSAR_OnClick()
  Dim Interface As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Integer

  If CurrentQuery.State <> 1 Then
    bsShowMessage("É preciso salvar antes de processar.", "E")
    CanContinue = False
    Exit Sub
  End If
  If VisibleMode Then
  	Set Interface = CreateBennerObject("BSINTERFACE0057.Processar")
  	Interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,vsMensagemRetorno)
  	Set Interface = Nothing

  	RefreshNodesWithTable("SAM_TAXASAUDESUPLEMENTAR")
  Else
	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "SamTaxaSaudeSuplementar", _
                                    "Processar", _
                                    "TSS - Taxa Saúde Suplementar - Processar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_TAXASAUDESUPLEMENTAR", _
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

  'Set Interface = CreateBennerObject("BSANS001.uROTINAS")
  'Interface.DuplicarModelo(CurrentSystem)
  'Set Interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()

    BOTAOCANCELAR.Visible = False
    BOTAOPROCESSAR.Visible = False
    COMPETENCIAINICIAL.ReadOnly = True
    COMPETENCIAFINAL.ReadOnly = True
    IDADEDIVISAO.ReadOnly = True
    VALORIDADESUP.ReadOnly = True
    VALORIDADEINF.ReadOnly = True
    DESCRICAO.ReadOnly = True

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And _
       (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime < CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime) Then
		bsShowMessage("A Competência final, se informada, deve ser maior ou igual a inicial.", "E")
  		CanContinue = False
	Else
  		CanContinue = True
	End If

	Dim Interface As Object
	Dim Linha As String
	Dim SQL As Object

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_TAXASAUDESUPLEMENTAR", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "OPERADORA", "")

	If Linha = "" Then
		CanContinue = True
	Else
	  	Set Interface = Nothing
	  	bsShowMessage(Linha, "E")
		CanContinue = False
	End If

	Set Interface = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
