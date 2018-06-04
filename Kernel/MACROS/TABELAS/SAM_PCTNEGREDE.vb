'HASH: 2DD95D79E909195C2C3D9FD57418CA0C
'Macro: SAM_PCTNEGREDE
'Mauricio Ibelli -14/08/2001 -sms3858 -Acesso nao tem campo usuario na tabela
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

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
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If

 '-------------VALOR TOTAL DO PACOTE ------------------------------------------------
  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then

    Dim interface As Object
    Dim valorpacte As Currency


    Set interface = CreateBennerObject("BSPRE001.Rotinas")
    valorpacte = interface.ValorTotalPacote(CurrentSystem, "SAM_PCTNEGREDE", CurrentQuery.FieldByName("HANDLE").Value)
    VALORPACOTE.Text = "Valor total do pacote: R$ " + Format(valorpacte, "#,##0.00")
  Else
    VALORPACOTE.Text = " "
  End If
  '-------------VALOR TOTAL DO PACOTE ------------------------------------------------

  If VisibleMode Then

	If CurrentQuery.FieldByName("evento").AsString = "" Then
    	Exit Sub
  	End If
  	GRAUAGERAR.LocalWhere = "ORIGEMVALOR =7 AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @EVENTO)
  Else
  	If WebMenuCode = "T2421" Then
  		PRESTADOR.ReadOnly = True
  	End If
	GRAUAGERAR.WebLocalWhere = "ORIGEMVALOR =7 AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO))
	If WebMenuCode = "T5674" Then
		EVENTO.ReadOnly = True
	End If
  End If

End Sub

'MILANI -SMS -22609

Public Sub BOTAOINCLUIRITENS_OnClick()
  If CurrentQuery.State <>1 Then
    If VisibleMode Then
      MsgBox("Os parâmetros não podem estar em edição")
    Else
      CancelDescription = "Os parâmetros não podem estar em edição"
    End If
    Exit Sub
  End If
  Dim interface As Object
  Set interface = CreateBennerObject("BSPRE009.ROTINAS")
  interface.ItensPacotes(CurrentSystem, "SAM_PCTNEGREDE_GRAU", CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing
End Sub

