'HASH: 2D6B017EBB146FEF4971A636B3497451

Public Sub Main
	Dim piHandleCompetencia As Long
	Dim psPgtoInicial As String
	Dim psPgtoFinal As String
	Dim psDigitadoAte As String
	Dim piPegInicial As Long
	Dim piPegFinal As Long
	Dim piHandleTipoPeg As Long
	Dim psSituacaoPeg As String
	Dim psRegime As String
	Dim psTipoGuia As String
	Dim psPegs As String
	Dim psMsgRetorno As String
	Dim piResult As Long
	Dim psCampoOrdem As String
	Dim SamPegDll As Object
    Dim pstipoPegNormalReapTodos As String
    Dim vsMsgRetorno As String
    Dim viResultado As Integer
    Dim psFilial As String


	piHandleCompetencia = CLng( ServiceVar("piHandleCompetencia") )
	psPgtoInicial = CStr( ServiceVar("psPgtoInicial") )
	psPgtoFinal = CStr( ServiceVar("psPgtoFinal") )
	psDigitadoAte = CStr( ServiceVar("psDigitadoAte") )
	piPegInicial = CLng( ServiceVar("piPegInicial") )
	piPegFinal = CLng( ServiceVar("piPegFinal") )
	piHandleTipoPeg = CLng( ServiceVar("piHandleTipoPeg") )
	psSituacaoPeg = CStr( ServiceVar("psSituacaoPeg") )
	psRegime = CStr( ServiceVar("psRegime") )
	psTipoGuia = CStr( ServiceVar("psTipoGuia") )
	psPegs = CStr( ServiceVar("psPegs") )
	psMsgRetorno = CStr( ServiceVar("psMsgRetorno") )
	piResult = CLng( ServiceVar("piResult") )
	pstipoPegNormalReapTodos = CStr(ServiceVar("pstipoPegNormalReapTodos"))
	psFilial = CStr(ServiceVar("psFilial"))

	psCampoOrdem = ""
    vsMsgRetorno = ""
    viResultado = 0

	On Error GoTo erro

	psPegs = CStr( ServiceVar("psPEGs") )


	Set SamPegDll = CreateBennerObject("SamPeg.PEGLote_Processar")
	psMsgRetorno = SamPegDll.BuscarPEGs(CurrentSystem, _
	   piHandleCompetencia, _
	   psPgtoInicial, _
	   psPgtoFinal, _
	   psDigitadoAte, _
	   piPegInicial, _
	   piPegFinal, _
	   piHandleTipoPeg, _
	   psSituacaoPeg, _
	   psRegime, _
	   psTipoGuia, _
	   psCampoOrdem, _
       pstipoPegNormalReapTodos, _
	   psPegs, _
	   psFilial)

    Set SamPegDll = Nothing

    piResult = 0
    If Trim(psMsgRetorno) <> "" Then
      piResult = 1
    End If

    GoTo fim

    erro:
      piResult = 1
      psMsgRetorno = Err.Description

    fim:
	ServiceVar("psPegs") = CStr(psPegs)
	ServiceVar("psMsgRetorno") = CStr(psMsgRetorno)
	ServiceVar("piResult") = CLng(piResult)

End Sub
