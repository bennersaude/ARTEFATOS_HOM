'HASH: 214E2E5DD6A8B33CEC2687DB90317E95
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim mensagemErro As String
	Dim retorno As Integer
	Dim IGerarDadosIR As Object


	vAnoCalendario = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger
	vGerarEmTabela = CInt(IIf(CurrentQuery.FieldByName("GERAREMTABELA").AsBoolean, 1, 0))
	vGerarEmArquivo = CInt(IIf(CurrentQuery.FieldByName("GERAREMARQUIVO").AsBoolean, 1, 0))


	If vGerarEmTabela = 0 And vGerarEmArquivo = 0 Then
		bsShowMessage("É necessário escolher pelo menos uma opção.","E")
		CanContinue = False
		Exit Sub
	End If
	Set IGerarDadosIR =CreateBennerObject("BSDMED.GERACAODADOSIR")
	retorno = IGerarDadosIR.ProcessarWeb(vAnoCalendario,vGerarEmTabela,vGerarEmArquivo,mensagemErro)
	If retorno = 1 Then
		bsShowMessage(mensagemErro, "I")
  	End If
  	Set IGerarDadosIR =Nothing
End Sub
