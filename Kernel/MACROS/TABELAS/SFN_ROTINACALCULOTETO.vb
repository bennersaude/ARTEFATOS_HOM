'HASH: DF78C17AA37022D926B8E7BEBFCD3374
'#Uses "*bsShowMessage"
Dim obj As Object
Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
	Dim obj As Object
	Dim HandleRotina As Integer
	Dim retorno As String
	HandleRotina = CurrentQuery.FieldByName("HANDLE").AsInteger
	If VisibleMode Then
		Set obj = CreateBennerObject("Benner.Saude.Financeiro.RotinaCalculoTetoMensal.CancelarRotinaCalculoTeto")
	   	retorno = obj.ExecDesk(HandleRotina)
		bsShowmessage(retorno, "I")
	Else
		CancelarRotina
	End If
	RefreshNodesWithTable("SFN_ROTINACALCULOTETO")
	Set obj = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
	Dim obj As Object
	Dim HandleRotina As Integer
	Dim retorno As String
	HandleRotina = CurrentQuery.FieldByName("HANDLE").AsInteger
	If VisibleMode Then
		Set obj = CreateBennerObject("Benner.Saude.Financeiro.RotinaCalculoTetoMensal.ProcessarRotinaCalculoTeto")
	   	retorno = obj.ExecDesk(HandleRotina)
		bsShowmessage(retorno, "I")
	Else
		ProcessarRotina
	End If
	RefreshNodesWithTable("SFN_ROTINACALCULOTETO")
	Set obj = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If ((CurrentQuery.FieldByName("VENCIMENTOINICIAL").IsNull And Not CurrentQuery.FieldByName("VENCIMENTOFINAL").IsNull ) Or (Not CurrentQuery.FieldByName("VENCIMENTOINICIAL").IsNull And CurrentQuery.FieldByName("VENCIMENTOFINAL").IsNull )) Then
		CanContinue = False
		bsShowMessage("É obrigatório o preenchimento das duas datas caso uma delas tenha sido preenchida.", "E")
		Exit Sub
	End If


	If(CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime > CurrentQuery.FieldByName("VENCIMENTOFINAL").AsDateTime) Then
		CanContinue = False
		bsShowMessage("A data de vencimento inicial não pode ser maior que a data de vencimento final.","E")
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	Select Case CommandID
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick

        Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick

        Case "BOTAO"
			BOTAOCANCELAR_OnClick

	End Select
End Sub

Public Sub ProcessarRotina()
	Dim vsMensagemErro As String
   	Dim viRetorno As Long
   	Dim HandleRotina As Integer
   	Dim CodigoRotina As Integer

  	HandleRotina = CurrentQuery.FieldByName("HANDLE").AsInteger
  	CodigoRotina = CurrentQuery.FieldByName("CODIGO").AsInteger

	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = obj.ExecucaoImediata(CurrentSystem, _
	                                "Benner.Saude.Financeiro.RotinaCalculoTetoMensal", _
	                                "ProcessarRotinaCalculoTeto", _
	                                "Processando Rotina Calculo Teto Mensal ("+CStr(CodigoRotina)+")", _
	                                HandleRotina, _
	                                "SFN_ROTINACALCULOTETO", _
	                                "SITUACAO", _
	                                "", _
	                                "", _
	                                "P", _
	                                False, _
	                                vsMensagemErro, _
	                                Null)

	If viRetorno = 0 Then
	 bsShowMessage("Processo enviado ao servidor, favor verificar o monitor de processos!", "I")
	Else
	 bsShowMessage(vsMensagemErro, "I")
	End If
	Set obj = Nothing
End Sub

Public Sub CancelarRotina()
	Dim vsMensagemErro As String
   	Dim viRetorno As Long
   	Dim HandleRotina As Integer
   	Dim CodigoRotina As Integer

  	HandleRotina = CurrentQuery.FieldByName("HANDLE").AsInteger
  	CodigoRotina = CurrentQuery.FieldByName("CODIGO").AsInteger

	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = obj.ExecucaoImediata(CurrentSystem, _
	                                "Benner.Saude.Financeiro.RotinaCalculoTetoMensal", _
	                                "CancelarRotinaCalculoTeto", _
	                                "Cancelando Rotina Calculo Teto Mensal ("+CStr(CodigoRotina)+")", _
	                                HandleRotina, _
	                                "SFN_ROTINACALCULOTETO", _
	                                "SITUACAO", _
	                                "", _
	                                "", _
	                                "C", _
	                                False, _
	                                vsMensagemErro, _
	                                Null)

	If viRetorno = 0 Then
	 bsShowMessage("Processo enviado ao servidor, favor verificar o monitor de processos!", "I")
	Else
	 bsShowMessage(vsMensagemErro, "I")
	End If
	Set obj = Nothing
End Sub
