'HASH: 4D824B51AB662646EF3CEFA89CBE6E30

Public Sub Main
	Dim vsCompetencias As String
	Dim vsCompetencia As String
	Dim vsTipoGuia As String
	Dim vsTipoPeg As String
	Dim vsDataPagamentoI As String
	Dim vsDataPagamentoF As String
	Dim vsMsgErro As String
	Dim piResult As Long

	Dim SamPegDll As Object
	Dim qAux As Object

    piResult = 0
    vsMsgErro = ""
	On Error GoTo erro

    '>>>>>>>>>>>>>>>>>> Buscando as competências, os tipos de guias e os tipos de PEG
	Set SamPegDll = CreateBennerObject("SamPeg.PEGLote_Processar")
	SamPegDll.InitInterface(CurrentSystem, _
	   vsCompetencias, _
	   vsTipoPeg, _
	   vsTipoGuia)

    Set SamPegDll = Nothing

    '>>>>>>>>>>>>>>>>>> Buscando a última data de pagamento existente no sistema
    Set qAux = NewQuery
    With qAux
      .Active = False
      .Clear
      .Add("SELECT MAX(DATAPAGAMENTO) DATAPAGAMENTO")
      .Add("   FROM SAM_PAGAMENTO")
      .Active = True
    End With

	vsDataPagamentoI = Format(Str(ServerDate), "DD/MM/YYYY")

    vsDataPagamentoF = ""
    If Not(qAux.FieldByName("DATAPAGAMENTO").IsNull) Then
		If qAux.FieldByName("DATAPAGAMENTO").AsDateTime < ServerDate Then
			vsDataPagamentoF = vsDataPagamentoI
		Else
        	vsDataPagamentoF = qAux.FieldByName("DATAPAGAMENTO").AsString
        End If
    End If

    '>>>>>>>>>>>>>>>>>> Buscando a última competência existente no sistema
    Set qAux = NewQuery
    With qAux
      .Active = False
      .Clear
      .Add("SELECT MAX(COMPETENCIA) COMPETENCIA")
      .Add("   FROM SAM_COMPETPEG")
      .Add(" WHERE COMPETENCIA >= :DATA")
      .ParamByName("DATA").AsDateTime = ServerDate
      .Active = True
    End With

    vsCompetencia = ""
    If Not(qAux.FieldByName("COMPETENCIA").IsNull) Then
        vsCompetencia = qAux.FieldByName("COMPETENCIA").AsString
    End If

    Set qAux = Nothing

    GoTo fim

	erro:
        piResult = 1
        vsMsgErro = Err.Description

    fim:

    '>>>>>>>>>>>>>>>>>> Retornando valores
	ServiceVar("psCompetencias") = CStr(vsCompetencias)
	ServiceVar("psTipoGuia") = CStr(vsTipoGuia)
	ServiceVar("psTipoPeg") = CStr(vsTipoPeg)

	ServiceVar("psCompetenciaAtual") = CStr(vsCompetencia)
    ServiceVar("psDataPagamentoI") = CStr(vsDataPagamentoI)
    ServiceVar("psDataPagamentoF") = CStr(vsDataPagamentoF)
    ServiceVar("psMsgRetorno") = CStr(vsMsgErro)
	ServiceVar("piResult") = CLng(piResult)

End Sub
