'HASH: 0A84A2834279AEDEB62D88CA80782A08
'Macro: SAM_PCTNEGFILIAL
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
  	ShowPopup = False
  	vHandle = ProcuraEvento(True, EVENTO.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("EVENTO").Value = vHandle
    	CurrentQuery.FieldByName("GRAUAGERAR").Clear
  	End If
End Sub

Public Sub TABLE_AfterScroll()
	'-------------VALOR TOTAL DO PACOTE ------------------------------------------------
  	If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    	Dim Interface As Object
    	Dim valorpacte As Currency

    	Set Interface = CreateBennerObject("BSPRE001.Rotinas")
     	valorpacte = Interface.ValorTotalPacote(CurrentSystem, "SAM_PCTNEGFILIAL", CurrentQuery.FieldByName("HANDLE").AsInteger)
    	VALORPACOTE.Text = "Valor total do pacote: R$ " + Format(valorpacte, "#,##0.00")
  	Else
    	VALORPACOTE.Text = " "
  	End If
  	'-------------VALOR TOTAL DO PACOTE ------------------------------------------------

  	If VisibleMode Then
  		If Not CurrentQuery.FieldByName("EVENTO").AsString = "" Then
    		GRAUAGERAR.LocalWhere = "ORIGEMVALOR = '7' AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @EVENTO)
  		End If
  	Else
  		GRAUAGERAR.WebLocalWhere = "ORIGEMVALOR = '7' AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO))
  		If WebMenuCode = "T5674" Then
  			FILIAL.ReadOnly = True
  		End If
  		If WebMenuCode = "T1303" Then
  			FILIAL.ReadOnly = True
  		End If
  	End If
End Sub

Public Sub BOTAOINCLUIRITENS_OnClick()
  	If CurrentQuery.State <>1 Then
    	bsShowMessage("Os parâmetros não podem estar em edição", "I")
    	Exit Sub
  	End If

  	Dim Interface As Object
  	Set Interface = CreateBennerObject("BSPRE009.ROTINAS")
  	Interface.ItensPacotes(CurrentSystem, "SAM_PCTNEGFILIAL_GRAU", CurrentQuery.FieldByName("HANDLE").AsInteger)
  	Set Interface = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  	If CommandID = "BOTAOINCLUIRITENS" Then
    	BOTAOINCLUIRITENS_OnClick
  	End If
End Sub
Public Sub BOTAOGERARRELATORIO_OnClick()
If CurrentQuery.State <> 1 Then
	    bsShowMessage("O registro está em edição.","I")
    Else

    	Dim RelatorioHandle As Long
		Dim QueryBuscaHandleRelatorio As Object


		Set QueryBuscaHandleRelatorio=NewQuery

		QueryBuscaHandleRelatorio.Add("SELECT RELATORIOPACOTE FROM SAM_PARAMETROSPRESTADOR")
    	        QueryBuscaHandleRelatorio.Active=False
   		QueryBuscaHandleRelatorio.Active=True
   		RelatorioHandle=QueryBuscaHandleRelatorio.FieldByName("RELATORIOPACOTE").AsInteger

		If (RelatorioHandle = 0) Then
		 bsShowMessage("Relatório não está parametrizado","I")
		 CanContinue = False
		Else
		 ReportPreview(RelatorioHandle,"", False, False)
		End If

	    Set QueryBuscaHandleRelatorio=Nothing
	End If
End Sub
