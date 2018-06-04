'HASH: F4DF865C3E6DA1623239FE0890A57EA8
'Macro: SAM_PRECOGENERICO_DOTAC
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String


	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOGENERICO_DOTAC", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "TABELAPRECO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()
	Dim PegDll As Object
	Dim qModeloGuia As BPesquisa
	Dim qParamGeral As BPesquisa
	Dim qSamTge As BPesquisa
	Dim qConvenio As BPesquisa
	Dim vlValorPrimeira As Double
	Dim vlValorSegunda As Double
	Dim vlValorDemais As Double
	Dim vlValorEvento As Currency
	Dim data As Date

	Set PegDll = CreateBennerObject("SAMPEG.Rotinas")
	Set qModeloGuia = NewQuery
	Set qParamGeral = NewQuery
	Set qSamTge = NewQuery
	Set qConvenio = NewQuery

	qModeloGuia.Clear
	qModeloGuia.Add("SELECT * FROM SAM_TIPOGUIA_MDGUIA")
	qModeloGuia.Active = True

	qParamGeral.Clear
	qParamGeral.Add("SELECT * FROM SAM_PARAMETROSPROCCONTAS, SAM_PARAMETROSATENDIMENTO")
	qParamGeral.Active = True

	qSamTge.Clear
	qSamTge.Add("SELECT GRAUPRINCIPAL FROM SAM_TGE WHERE HANDLE = :HANDLE")
	qSamTge.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
	qSamTge.Active = True

	qConvenio.Clear
	qConvenio.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
	qConvenio.Active = True

    Dim vDataBaseChecagemVigencia As Date

' Paulo Melo - SMS 118697 - 01/10/2009 - Inicio
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > ServerDate Then
    vDataBaseChecagemVigencia = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  Else
	If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
	  vDataBaseChecagemVigencia = ServerDate
	Else
	  vDataBaseChecagemVigencia = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
	End If
  End If
' Paulo Melo - SMS 118697 - 01/10/2009 - Fim

    Dim Interface As Object
    Dim ValorEvento As Currency
    LABELPRECO.Text = "" 
    If CurrentQuery.FieldByName("EVENTO").IsNull Then
      LABELPRECO.Text = ""
    Else
      Set Interface = CreateBennerObject("BSPRE001.Rotinas")
      ValorEvento = Interface.ValorEvento(CurrentSystem,  vDataBaseChecagemVigencia, 0, -1, -1, -1, -1, -1, -1, -1, CurrentQuery.FieldByName("TABELAPRECO").Value, CurrentQuery.FieldByName("EVENTO").Value, -1, -1, 0, "", -1)
      LABELPRECO.Text = "Valor do evento nesta vigência: R$ " + Format(ValorEvento, "#,##0.0000")+" ("+Format(ValorEvento, "#,##0.00")+")"
    End If
    Set Interface = Nothing
End Sub
