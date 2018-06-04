'HASH: B4D1C3B70D04F2BD0C69922EACB4BE47
'Macro: SAM_PRECOGERAL_FX
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaGenerica"
'#Uses "*bsShowMessage"

Option Explicit


Public Sub EVENTOFINAL_OnEnter()
	EVENTOFINAL.AnyLevel = True
End Sub

Public Sub EVENTOFINAL_OnExit()
	TABELAUS.SetFocus
End Sub

Public Sub EVENTOINICIAL_OnEnter()
	EVENTOINICIAL.AnyLevel = True
End Sub

Public Sub EVENTOINICIAL_OnExit()
	EVENTOFINAL.SetFocus
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim Interface As Object
	Dim vColunas, vCriterio, vCampos, vTabela As String
	Dim ProcuraEvento As Long
	Set Interface = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

	vCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString
	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vCriterio = vCriterio + criarCriterio
	End If

	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos  ", True, "")

	If vHandle > 0 Then
		CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
	End If

	Set Interface = Nothing
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim Interface As Object
	Dim vColunas, vCriterio, vCampos, vTabela As String
	Dim ProcuraEvento As Long
	Set Interface = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Evento Inicial antes de selecionar Evento Final", "E")
	  Exit Sub
	End If

	vCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString
	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vCriterio = vCriterio + criarCriterio
	End If

	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos ", True, "")

	If vHandle > 0 Then
		CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
	End If

	Set Interface = Nothing
End Sub

Public Sub MASCARATGE_OnChange()
	Dim qAux As Object
	Set qAux = NewQuery
	qAux.Clear
	qAux.Add("SELECT MASCARATGE FROM SAM_TGE WHERE HANDLE = :HANDLE")
	qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
	qAux.Active = True
	If qAux.FieldByName("MASCARATGE").AsInteger <> CurrentQuery.FieldByName("MASCARATGE").AsInteger Then
	  CurrentQuery.FieldByName("EVENTOINICIAL").Clear
	  CurrentQuery.FieldByName("ESTRUTURAINICIAL").Clear
	End If
	qAux.Clear
	qAux.Add("SELECT MASCARATGE FROM SAM_TGE WHERE HANDLE = :HANDLE")
	qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
	qAux.Active = True
	If qAux.FieldByName("MASCARATGE").AsInteger <> CurrentQuery.FieldByName("MASCARATGE").AsInteger Then
	  CurrentQuery.FieldByName("EVENTOFINAL").Clear
	  CurrentQuery.FieldByName("ESTRUTURAFINAL").Clear
	End If
	qAux.Active = True
End Sub

Public Sub TABLE_AfterEdit()
  If (CurrentQuery.State = 2) And (CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0) Then
    Dim qAux As Object
    Dim vIMascaraTGEInicial As Integer
    Dim vIMascaraTGEFinal As Integer
    Set qAux = NewQuery
    qAux.Clear
    qAux.Add("SELECT MASCARATGE FROM SAM_TGE WHERE HANDLE = :HANDLE")
    qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
    qAux.Active = True
    vIMascaraTGEInicial = qAux.FieldByName("MASCARATGE").AsInteger

    qAux.Clear
    qAux.Add("SELECT MASCARATGE FROM SAM_TGE WHERE HANDLE = :HANDLE")
    qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
    qAux.Active = True
    vIMascaraTGEFinal = qAux.FieldByName("MASCARATGE").AsInteger
    Set qAux = Nothing
    If vIMascaraTGEInicial = vIMascaraTGEFinal Then
      CurrentQuery.FieldByName("MASCARATGE").AsInteger = vIMascaraTGEInicial
    Else
      If CurrentQuery.FieldByName("MASCARATGE").AsInteger > 0 Then
        bsShowMessage("Evento Inicial e Evento Final não fazem parte da mesma máscara", "E")
	    Exit Sub
	  End If
    End If

  End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T4377" Then
			CONVENIO.ReadOnly = True
		End If

		EVENTOINICIAL.WebLocalWhere = criarCriterio
		EVENTOFINAL.WebLocalWhere = criarCriterio
	End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		CanContinue = False
		bsShowMessage("Registro finalizado não pode ser alterado", "E")
		Exit Sub
	End If
End Sub

Public Sub TABELAFILME_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaFilme(TABELAFILME.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAFILME").Value = vHandle
	End If
End Sub

Public Sub TABELAPRECO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaGenerica(TABELAPRECO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAPRECO").Value = vHandle
	End If
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaUS(TABELAUS.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAUS").Value = vHandle
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Dim EstruturaI As String
	Dim EstruturaF As String

    Dim qVerifica As Object
    Set qVerifica = NewQuery

    qVerifica.Add("SELECT MASCARATGE")
    qVerifica.Add("  FROM SAM_TGE")
	qVerifica.Add(" WHERE HANDLE = :HANDLE")

	qVerifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
    qVerifica.Active = True

	If qVerifica.FieldByName("MASCARATGE").AsInteger <> CurrentQuery.FieldByName("MASCARATGE").AsInteger Then
		CanContinue = False
		bsShowMessage("O Evento Inicial possui máscara diferente da informada. Altere o evento ou a máscara.", "E")
		Exit Sub
	End If

    qVerifica.Active = False
	qVerifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
	qVerifica.Active = True

	If qVerifica.FieldByName("MASCARATGE").AsInteger <> CurrentQuery.FieldByName("MASCARATGE").AsInteger Then
		CanContinue = False
		bsShowMessage("O Evento Final possui máscara diferente da informada. Altere o evento ou a máscara.", "E")
		Exit Sub
	End If

	Dim SQLTGE, SQLMASC As Object
	Dim Estrutura As String
	Set SQLTGE = NewQuery

	SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")

	SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
	SQLTGE.Active = True

	CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = SQLTGE.FieldByName("ESTRUTURA").Value

	SQLTGE.Active = False
	SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
	SQLTGE.Active = True

	Estrutura = SQLTGE.FieldByName("ESTRUTURA").Value

	SQLTGE.Active = False

	CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Estrutura

	Set SQLTGE = Nothing



	If CanContinue = True Then
		EstruturaI = CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
		EstruturaF = CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString

		Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

		If CurrentQuery.FieldByName("CONVENIO").IsNull Then
			Condicao = Condicao + " CONVENIO IS NULL"
		Else
			Condicao = Condicao + " CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
		End If

		If CurrentQuery.FieldByName("MASCARATGE").AsInteger > 0 Then
			Condicao = Condicao + " AND MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString
		End If


		Linha = Interface.EventoFx(CurrentSystem, "SAM_PRECOGERAL_FX", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao)

		If Linha = "" Then
			CanContinue = True
		Else
			CanContinue = False
			bsShowMessage(Linha, "E")
			Exit Sub
		End If

		Set Interface = Nothing
	End If

	If CanContinue = True Then
		CanContinue = CheckEventosFx
	End If
End Sub

Public Function CheckEventosFx As Boolean
	CheckEventosFx = True

	If Not CurrentQuery.FieldByName("EVENTOINICIAL").IsNull Then
		If CurrentQuery.FieldByName("EVENTOFINAL").IsNull Then
			CurrentQuery.FieldByName("EVENTOFINAL").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
		Else
			If CurrentQuery.FieldByName("EVENTOINICIAL").Value <>CurrentQuery.FieldByName("EVENTOFINAL").Value Then
				Dim SQLI, SQLF As Object
				Set SQLI = NewQuery

				SQLI.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOI")

				SQLI.ParamByName("HEVENTOI").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
				SQLI.Active = True

				Set SQLF = NewQuery

				SQLF.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOF")

				SQLF.ParamByName("HEVENTOF").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
				SQLF.Active = True

				If SQLF.FieldByName("ESTRUTURA").Value <SQLI.FieldByName("ESTRUTURA").Value Then
					bsShowMessage("Evento final não pode ser menor que o evento inicial!", "E")
					EVENTOFINAL.SetFocus
					CheckEventosFx = False
				End If

				Set SQLI = Nothing
				Set SQLF = Nothing
			End If
		End If
	End If
End Function

Public Sub BOTAOVALORES_OnClick()
    If CurrentQuery.State <> 1 Then ' consulta
        Dim vTipo As String
        vTipo = "edição"
        If CurrentQuery.State = 3 Then ' inserção
           vTipo = "inserção"
    	End If
        bsShowMessage("Registro em " + vTipo + "! Impossibilitado de executar a funcionalidade de consulta de preço.", "I")
    	Exit Sub
    End If

	Dim Interface As Object

	If Not CurrentQuery.FieldByName("TABELAPRECO").IsNull Then
		Set Interface = CreateBennerObject("BSPRE001.Rotinas")

		Interface.ValoresFxEventos(CurrentSystem, 10, -1, -1, CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value, CurrentQuery.FieldByName("ESTRUTURAFINAL").Value, -1, CurrentQuery.FieldByName("CONVENIO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("MASCARATGE").AsInteger)
	Else
		bsShowMessage("Para visualizar os preços dos eventos desta configuração de faixa, a tabela genérica deve ser informada !", "I")
	End If
End Sub

Public Sub TABLE_AfterInsert()
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

	SQL.Active = True

	If SQL.FieldByName("TOTAL").AsInteger = 1 Then
		SQL.Active = False
		SQL.Clear
		SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
		SQL.Active = True

		CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
	End If

	Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOVALORES"
			BOTAOVALORES_OnClick
	End Select
End Sub

Public Function criarCriterio As String
  If VisibleMode Then
    criarCriterio = criarCriterio _
                    +  " AND EXISTS (SELECT 1" _
	  	            + "               FROM SAM_PRECOGENERICO_DOTAC X" _
		            + "              WHERE X.TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString _
		            + "                AND X.EVENTO = SAM_TGE.HANDLE)"
  Else
    criarCriterio =   " (@CAMPO(MASCARATGE) = A.MASCARATGE) " _
                    + " AND ((@~CAMPO(TABELAPRECO) = -1)    " _
                    + "       OR EXISTS (SELECT 1" _
	  	            + "             FROM SAM_PRECOGENERICO_DOTAC X" _
		            + "            WHERE X.TABELAPRECO = @~CAMPO(TABELAPRECO)" _
		            + "              AND X.EVENTO = A.HANDLE)" _
		            + "      )"

  End If
End Function
