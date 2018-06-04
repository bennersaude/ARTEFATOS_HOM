'HASH: 4BB8A1E69C04800486DA8C7146EC719E
'Macro: SAM_PRECOESTADOREGIME_FX
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
	ShowPopup = False

	Dim Interface  As Object
	Dim vsData     As String
	Dim vsColunas  As String
	Dim vsCriterio As String
	Dim vsCampos   As String
	Dim vsTabela   As String
	Dim viHandle   As Long

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

	vsTabela  = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vsCampos  = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vsColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"
	vsCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vsCriterio = vsCriterio + criarCriterio
	End If

	Set Interface = CreateBennerObject("Procura.Procurar")

	viHandle = Interface.Exec(CurrentSystem, vsTabela, vsColunas, 1, vsCampos, vsCriterio, "Tabela Geral de Eventos", True, EVENTOINICIAL.Text)

	If (viHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = viHandle
	End If

	Set Interface = Nothing
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
	ShowPopup = False

	Dim Interface  As Object
	Dim vsData     As String
	Dim vsColunas  As String
	Dim vsCriterio As String
	Dim vsCampos   As String
	Dim vsTabela   As String
	Dim viHandle   As Long

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Evento Inicial antes de selecionar Evento Final", "E")
	  Exit Sub
	End If

	vsTabela  = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vsCampos  = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vsColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"
	vsCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vsCriterio = vsCriterio + criarCriterio
	End If

	Set Interface = CreateBennerObject("Procura.Procurar")

	viHandle = Interface.Exec(CurrentSystem, vsTabela, vsColunas, 1, vsCampos, vsCriterio, "Tabela Geral de Eventos", True, EVENTOFINAL.Text)

	If (viHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTOFINAL").AsInteger = viHandle
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

Public Sub TABLE_AfterScroll()
  If WebMode Then
    EVENTOINICIAL.WebLocalWhere = criarCriterio
    EVENTOFINAL.WebLocalWhere   = criarCriterio
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "E", CurrentQuery.FieldByName("ESTADO").AsInteger, "E") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode excluir", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "E", CurrentQuery.FieldByName("ESTADO").AsInteger, "A") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode alterar", "E")
		CanContinue = False
		Exit Sub
	End If

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		CanContinue = False
		bsShowMessage("Registro finalizado não pode ser alterado", "E")
		Exit Sub
	End If
End Sub

'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*ProcuraTabelaGenerica"
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
		CurrentQuery.FieldByName("EVENTOINICIAL").Value = Null
		CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = Null
		CurrentQuery.FieldByName("EVENTOFINAL").Value = Null
		CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Null
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

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "E", RecordHandleOfTable("ESTADOS"), "I") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode incluir", "E")
		CanContinue = False
		Exit Sub
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

	' Atribuir ESTRUTURAINICIAL E FINAL
	Dim SQLTGE, SQLMASC As Object
	Dim Estrutura As String
	' Atribuir ESTRUTURAINICIAL
	Set SQLTGE = NewQuery

	SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")

	SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
	SQLTGE.Active = True

	CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = SQLTGE.FieldByName("ESTRUTURA").Value

	' Atribuir ESTRUTURAFINAL
	SQLTGE.Active = False
	SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
	SQLTGE.Active = True

	Estrutura = SQLTGE.FieldByName("ESTRUTURA").Value

	SQLTGE.Active = False

	Set SQLTGE = Nothing
	' Completar ESTRUTURAFINAL com 99999
	Set SQLMASC = NewQuery

	SQLMASC.Add("SELECT M.MASCARA MASCARA FROM Z_TABELAS T, Z_MASCARAS M")
	SQLMASC.Add("WHERE T.NOME = 'SAM_TGE' AND M.TABELA = T.HANDLE")

	SQLMASC.Active = True

	Estrutura = Estrutura + Mid(SQLMASC.FieldByName("MASCARA").AsString, Len(Estrutura) + 1)

	CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Estrutura

	SQLMASC.Active = False

	Set SQLMASC = Nothing

	' Checar Vigencia
	Condicao = " ESTADO = " + CurrentQuery.FieldByName("ESTADO").AsString
	Condicao = Condicao + " AND REGIMEATENDIMENTO = " + CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsString

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger > 0 Then
    	Condicao = Condicao + " AND MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString
	End If

	EstruturaI = CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
	EstruturaF = CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.EventoFx(CurrentSystem, "SAM_PRECOESTADOREGIME_FX", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing

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
	Dim Interface As Object
	Dim SQL As Object
	Dim Nivel As Integer
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active = True

	Nivel = -1

	If SQL.FieldByName("NIVEL1").AsInteger = 7 Then
		Nivel = 1
	ElseIf SQL.FieldByName("NIVEL2").AsInteger = 7 Then
		Nivel = 2
	ElseIf SQL.FieldByName("NIVEL3").AsInteger = 7 Then
		Nivel = 3
	ElseIf SQL.FieldByName("NIVEL4").AsInteger = 7 Then
		Nivel = 4
	ElseIf SQL.FieldByName("NIVEL5").AsInteger = 7 Then
		Nivel = 5
	ElseIf SQL.FieldByName("NIVEL6").AsInteger = 7 Then
		Nivel = 6
	ElseIf SQL.FieldByName("NIVEL7").AsInteger = 7 Then
		Nivel = 7
	ElseIf SQL.FieldByName("NIVEL8").AsInteger = 7 Then
		Nivel = 8
	End If

	If Nivel <> -1 Then
		If Not CurrentQuery.FieldByName("TABELAPRECO").IsNull Then
			Set Interface = CreateBennerObject("BSPRE001.Rotinas")

			Interface.ValoresFxEventos(CurrentSystem, 20, CurrentQuery.FieldByName("ESTADO").Value, -1, CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value, CurrentQuery.FieldByName("ESTRUTURAFINAL").Value, CurrentQuery.FieldByName("REGIMEATENDIMENTO").Value, CurrentQuery.FieldByName("CONVENIO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("MASCARATGE").AsInteger)
		Else
			bsShowMessage("Para visualizar os preços dos eventos desta configuração de faixa, a tabela genérica deve ser informada !", "I")
		End If
	Else
		bsShowMessage("Na configuração de busca do preço, não foi definido um nível para o Estado !", "I")
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

Public Sub EVENTOFINAL_OnExit()
	TABELAUS.SetFocus
End Sub

Public Sub EVENTOINICIAL_OnExit()
	EVENTOFINAL.SetFocus
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
                    + " AND SAM_TGE.HANDLE IN (SELECT EVENTO" _
	  	            + "                          FROM SAM_PRECOGENERICO_DOTAC" _
		            + "                         WHERE TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString + ")"
  Else
    criarCriterio = " (@CAMPO(MASCARATGE) = A.MASCARATGE) " _
    				+ " AND ((@~CAMPO(TABELAPRECO) = -1) " _
                    + "      OR A.HANDLE IN (SELECT EVENTO" _
	  	            + "                        FROM SAM_PRECOGENERICO_DOTAC" _
		            + "                       WHERE TABELAPRECO = @~CAMPO(TABELAPRECO))" _
		            + "      )"

  End If
End Function
