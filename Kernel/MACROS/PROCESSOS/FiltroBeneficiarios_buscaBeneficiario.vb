'HASH: 3A4852B6259602E687FBA209C3439A95

Public Sub Main
	On Error GoTo erro

	Dim pDataAtendimento As Date
	Dim pBeneficiario As String
	Dim pContrato As Long
	Dim pEstado As Long
	Dim pMunicipio As Long
	Dim pCPF As String
	Dim pRG As String
	Dim pRecebedor As Long
	Dim pAtivos As Boolean
	Dim pInativos As Boolean
	Dim pTodos As Boolean
	Dim pBeneficiarioTitular As Long
	Dim pNome As String
	Dim pAntigo As String
	Dim pAfinidade As String
	Dim pCartao As String
	Dim pPosicao As Long
	Dim pCodOrigem As String
	Dim pCodRepasse As String
	Dim pMatFuncional As String
	Dim pRetornoDados As String
	Dim pMensagem As String
	Dim pResultado As Long

	Dim CA010Dll As Object

	pDataAtendimento = CDate( ServiceVar("pDataAtendimento") )
	pBeneficiario = ( CStr( ServiceVar("pBeneficiario") ) )
	pContrato = CLng( ServiceVar("pContrato") )
	pEstado = CLng( ServiceVar("pEstado") )
	pMunicipio = CLng( ServiceVar("pMunicipio") )
	pCPF = ( CStr( ServiceVar("pCPF") ))
	pRG = ( CStr( ServiceVar("pRG") ) )
	pRecebedor = CLng( ServiceVar("pRecebedor") )
	pAtivos = CBool( ServiceVar("pAtivos") )
	pInativos = CBool( ServiceVar("pInativos") )
	pTodos = CBool( ServiceVar("pTodos") )
	pBeneficiarioTitular = CLng( ServiceVar("pBeneficiarioTitular") )
	pNome = ( CStr( ServiceVar("pNome") ) )
	pAntigo = ( CStr( ServiceVar("pAntigo") ) )
	pAfinidade = ( CStr( ServiceVar("pAfinidade") ) )
	pCartao = ( CStr( ServiceVar("pCartao") ) )
	pPosicao = CLng( ServiceVar("pPosicao") )
	pCodOrigem = ( CStr( ServiceVar("pCodOrigem") ) )
	pCodRepasse = ( CStr( ServiceVar("pCodRepasse") ) )
	pMatFuncional = ( CStr( ServiceVar("pMatFuncional") ) )
	'pRetornoDados = ( CStr( ServiceVar("pRetornoDados") ) )
	'pMensagem = ( CStr( ServiceVar("pMensagem") ) )
	'pResultado = CLng( ServiceVar("pResultado") )


   If pBeneficiarioTitular = 1 Then
   		Dim SQL As Object
   		Set SQL = NewQuery
   		SQL.Add("SELECT BENEFICIARIO FROM SAM_PEG WHERE HANDLE = :HANDLE")
   		SQL.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HPEG"))
   		SQL.Active = True

        If SQL.FieldByName("BENEFICIARIO").AsInteger > 0 Then
        	pBeneficiarioTitular = SQL.FieldByName("BENEFICIARIO").AsInteger
        Else
        	pBeneficiarioTitular = 0
		End If
   		Set SQL = Nothing
  End If


	Set CA010Dll = CreateBennerObject("CA010.ConsultaBeneficiario")
	pResultado = CA010Dll.SelecionaBeneficiario(CurrentSystem, _
	                                                                       pDataAtendimento, _
	                                                                       pBeneficiario, _
	                                                                       pContrato, _
	                                                                       pEstado, _
	                                                                       pMunicipio, _
	                                                                       pCPF, _
	                                                                       pRG, _
	                                                                       pRecebedor, _
	                                                                       pAtivos, _
	                                                                       pInativos, _
	                                                                       pTodos, _
	                                                                       pBeneficiarioTitular, _
	                                                                       pNome, _
	                                                                       pAntigo, _
	                                                                       pAfinidade, _
	                                                                       pCartao, _
	                                                                       pPosicao, _
	                                                                       pCodOrigem, _
	                                                                       pCodRepasse, _
	                                                                       pMatFuncional, _
	                                                                       pRetornoDados, _
	                                                                       pMensagem)

    Set CA010Dll = Nothing
	ServiceVar("pRetornoDados") = ( (pRetornoDados) )

	ServiceVar("pMensagem") = ( CStr(pMensagem) )

	ServiceVar("pResultado") = CLng(pResultado)

    Exit Sub
    erro:
    	Set CA010Dll = Nothing
		psMensagem = Err.Description
		ServiceVar("pResultado") = 1
		ServiceVar("pMensagem") = CStr( psMensagem )

End Sub
