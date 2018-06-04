'HASH: DA5430DEA35D1F34E8BEE25678237F16
'Macro: SAM_PRECOLOCALEXEC_FAIXA
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*ProcuraTabelaGenerica"

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
	Dim qPrestador As Object
	Set qPrestador = NewQuery

	qPrestador.Add("SELECT A.PRESTADOR,                                                ")
	qPrestador.Add("       B.ASSOCIACAO                                                ")
	qPrestador.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID A                            ")
	qPrestador.Add("  JOIN SAM_PRESTADOR                  B ON (B.HANDLE = A.PRESTADOR)")
	qPrestador.Add(" WHERE A.HANDLE = :HANDLE                                          ")

	qPrestador.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CORPOCLINICO").AsInteger
	qPrestador.Active = True

	vsData = SQLDate(ServerDate)
	vsCriterio = ""

	If (qPrestador.FieldByName("ASSOCIACAO").AsString <> "S") Then
		vsCriterio = vsCriterio + "((SAM_TGE.HANDLE IN (SELECT DISTINCT GE.EVENTO "
		vsCriterio = vsCriterio + "                       FROM SAM_ESPECIALIDADEGRUPO_EXEC         GE "
		vsCriterio = vsCriterio + "                       JOIN SAM_ESPECIALIDADEGRUPO              EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO) "
		vsCriterio = vsCriterio + "                       JOIN SAM_ESPECIALIDADE                   E  ON (E.HANDLE = EG.ESPECIALIDADE) "
		vsCriterio = vsCriterio + "                       JOIN SAM_PRESTADOR_ESPECIALIDADE         PE ON (PE.ESPECIALIDADE = E.HANDLE) "
		vsCriterio = vsCriterio + "                       LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE) "
		vsCriterio = vsCriterio + "                      WHERE PE.DATAINICIAL <= " + vsData
		vsCriterio = vsCriterio + "                        AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= " + vsData + ") "
		vsCriterio = vsCriterio + "                        AND PE.PRESTADOR = " + qPrestador.FieldByName("PRESTADOR").AsString
		vsCriterio = vsCriterio + "                        AND GE.EVENTO NOT IN (SELECT X.EVENTO "
		vsCriterio = vsCriterio + "                                                FROM SAM_PRESTADOR_REGRA X "
		vsCriterio = vsCriterio + "                                               WHERE X.REGRAEXCECAO   = 'E' "
		vsCriterio = vsCriterio + "                                                 AND X.PERMITERECEBER = 'S' "
		vsCriterio = vsCriterio + "                                                 AND X.PRESTADOR      = PE.PRESTADOR "
		vsCriterio = vsCriterio + "                                                 AND X.DATAINICIAL   <= " + vsData
		vsCriterio = vsCriterio + "                                                 AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vsData + ")))) OR "
		vsCriterio = vsCriterio + " (SAM_TGE.HANDLE IN(SELECT X.EVENTO "
		vsCriterio = vsCriterio + "                      FROM SAM_PRESTADOR_REGRA X "
		vsCriterio = vsCriterio + "                     WHERE X.REGRAEXCECAO   = 'R' "
		vsCriterio = vsCriterio + "                       AND X.PERMITERECEBER = 'S' "
		vsCriterio = vsCriterio + "                       AND X.PRESTADOR      = " + qPrestador.FieldByName("PRESTADOR").AsString
		vsCriterio = vsCriterio + "                       AND X.DATAINICIAL   <= " + vsData
		vsCriterio = vsCriterio + "                       AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vsData + ")))) "
	Else
		vsCriterio = vsCriterio + "SAM_TGE.ULTIMONIVEL = 'S'"
	End If

	Set qPrestador = Nothing

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vsCriterio = vsCriterio + "AND (SAM_TGE.HANDLE IN (SELECT EVENTO
		vsCriterio = vsCriterio + "                          FROM SAM_PRECOGENERICO_DOTAC
		vsCriterio = vsCriterio + "                         WHERE TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString + ")) "
	End If

	vsTabela  = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vsCampos  = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vsColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

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
	Dim qPrestador As Object
	Set qPrestador = NewQuery

	qPrestador.Add("SELECT A.PRESTADOR,                                                ")
	qPrestador.Add("       B.ASSOCIACAO                                                ")
	qPrestador.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID A                            ")
	qPrestador.Add("  JOIN SAM_PRESTADOR                  B ON (B.HANDLE = A.PRESTADOR)")
	qPrestador.Add(" WHERE A.HANDLE = :HANDLE                                          ")

	qPrestador.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CORPOCLINICO").AsInteger
	qPrestador.Active = True

	vsData = SQLDate(ServerDate)
	vsCriterio = ""

	If (qPrestador.FieldByName("ASSOCIACAO").AsString <> "S") Then
		vsCriterio = vsCriterio + "((SAM_TGE.HANDLE IN (SELECT DISTINCT GE.EVENTO "
		vsCriterio = vsCriterio + "                       FROM SAM_ESPECIALIDADEGRUPO_EXEC         GE "
		vsCriterio = vsCriterio + "                       JOIN SAM_ESPECIALIDADEGRUPO              EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO) "
		vsCriterio = vsCriterio + "                       JOIN SAM_ESPECIALIDADE                   E  ON (E.HANDLE = EG.ESPECIALIDADE) "
		vsCriterio = vsCriterio + "                       JOIN SAM_PRESTADOR_ESPECIALIDADE         PE ON (PE.ESPECIALIDADE = E.HANDLE) "
		vsCriterio = vsCriterio + "                       LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE) "
		vsCriterio = vsCriterio + "                      WHERE PE.DATAINICIAL <= " + vsData
		vsCriterio = vsCriterio + "                        AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= " + vsData + ") "
		vsCriterio = vsCriterio + "                        AND PE.PRESTADOR = " + qPrestador.FieldByName("PRESTADOR").AsString
		vsCriterio = vsCriterio + "                        AND GE.EVENTO NOT IN (SELECT X.EVENTO "
		vsCriterio = vsCriterio + "                                                FROM SAM_PRESTADOR_REGRA X "
		vsCriterio = vsCriterio + "                                               WHERE X.REGRAEXCECAO   = 'E' "
		vsCriterio = vsCriterio + "                                                 AND X.PERMITERECEBER = 'S' "
		vsCriterio = vsCriterio + "                                                 AND X.PRESTADOR      = PE.PRESTADOR "
		vsCriterio = vsCriterio + "                                                 AND X.DATAINICIAL   <= " + vsData
		vsCriterio = vsCriterio + "                                                 AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vsData + ")))) OR "
		vsCriterio = vsCriterio + " (SAM_TGE.HANDLE IN(SELECT X.EVENTO "
		vsCriterio = vsCriterio + "                      FROM SAM_PRESTADOR_REGRA X "
		vsCriterio = vsCriterio + "                     WHERE X.REGRAEXCECAO   = 'R' "
		vsCriterio = vsCriterio + "                       AND X.PERMITERECEBER = 'S' "
		vsCriterio = vsCriterio + "                       AND X.PRESTADOR      = " + qPrestador.FieldByName("PRESTADOR").AsString
		vsCriterio = vsCriterio + "                       AND X.DATAINICIAL   <= " + vsData
		vsCriterio = vsCriterio + "                       AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vsData + ")))) "
	Else
		vsCriterio = vsCriterio + "SAM_TGE.ULTIMONIVEL = 'S' "
	End If

	Set qPrestador = Nothing

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vsCriterio = vsCriterio + "AND (SAM_TGE.HANDLE IN (SELECT EVENTO
		vsCriterio = vsCriterio + "                          FROM SAM_PRECOGENERICO_DOTAC
		vsCriterio = vsCriterio + "                         WHERE TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString + ")) "
	End If

	vsTabela  = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vsCampos  = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vsColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	Set Interface = CreateBennerObject("Procura.Procurar")

	viHandle = Interface.Exec(CurrentSystem, vsTabela, vsColunas, 1, vsCampos, vsCriterio, "Tabela Geral de Eventos", True, EVENTOFINAL.Text)

	If (viHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTOFINAL").AsInteger = viHandle
	End If

	Set Interface = Nothing
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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim INTERFACE As Object
	Dim Linha As String
	Dim Condicao As String
	Dim EstruturaI As String
	Dim EstruturaF As String
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
	' Completar ESTRUTURAFinal com 99999
	Set SQLMASC = NewQuery

	SQLMASC.Add("SELECT M.MASCARA MASCARA FROM Z_TABELAS T, Z_MASCARAS M")
	SQLMASC.Add("WHERE T.NOME = 'SAM_TGE' AND M.TABELA = T.HANDLE")

	SQLMASC.Active = True

	Estrutura = Estrutura + Mid(SQLMASC.FieldByName("MASCARA").AsString, Len(Estrutura) + 1)

	CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Estrutura

	SQLMASC.Active = False

	Set SQLMASC = Nothing


	If CanContinue = True Then
		' Checar Vigencia
		EstruturaI = CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
		EstruturaF = CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString
		Condicao = " CORPOCLINICO = " + CurrentQuery.FieldByName("CORPOCLINICO").AsString

		If CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull Then
			Condicao = Condicao + " AND REGIMEATENDIMENTO IS NULL"
		Else
			Condicao = Condicao + " AND REGIMEATENDIMENTO = " + CurrentQuery.FieldByName("CORPOCLINICO").AsString
		End If

		If CurrentQuery.FieldByName("CONVENIO").IsNull Then
			Condicao = Condicao + " AND CONVENIO IS NULL"
		Else
			Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
		End If

		If CurrentQuery.FieldByName("CONVENIO").IsNull Then
			Condicao = Condicao + " AND CONVENIO IS NULL"
		Else
			Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
		End If

		Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")

		Linha = INTERFACE.EventoFx(CurrentSystem, "SAM_PRECOLOCALEXEC_FAIXA", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao)

		If Linha = "" Then
			CanContinue = True
		Else
			CanContinue = False
			bsShowMessage(Linha, "E")
			Exit Sub
		End If

		Set INTERFACE = Nothing
	End If

	If CanContinue = True Then
		CanContinue = CheckEventosFx
	End If

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Clear

	Dim vPrestador As String

	SQL.Clear

	SQL.Add("SELECT PRESTADOR FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE HANDLE = :HANDLE")

	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CORPOCLINICO").AsInteger
	SQL.Active = True

	vPrestador = SQL.FieldByName("PRESTADOR").AsString

	SQL.Active = False

	SQL.Clear

	SQL.Add("SELECT PRESTADOR FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE HANDLE = :HANDLE")

	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CORPOCLINICO").AsInteger
	SQL.Active = True

	vPrestador = SQL.FieldByName("PRESTADOR").AsString

	SQL.Active = False

	SQL.Clear

	Set SQL = Nothing
End Sub

Public Function CheckEventosFx As Boolean
	CheckEventosFx = True

	If Not CurrentQuery.FieldByName("EVENTOINICIAL").IsNull Then
		If CurrentQuery.FieldByName("EVENTOFINAL").IsNull Then
			CurrentQuery.FieldByName("EVENTOFINAL").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
		Else
			If CurrentQuery.FieldByName("EVENTOINICIAL").Value <> CurrentQuery.FieldByName("EVENTOFINAL").Value Then
				Dim SQLI, SQLF As Object
				Set SQLI = NewQuery

				SQLI.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOI")

				SQLI.ParamByName("HEVENTOI").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
				SQLI.Active = True

				Set SQLF = NewQuery

				SQLF.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOF")

				SQLF.ParamByName("HEVENTOF").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
				SQLF.Active = True

				If SQLF.FieldByName("ESTRUTURA").Value < SQLI.FieldByName("ESTRUTURA").Value Then
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
	Dim INTERFACE As Object
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active = True

	If (SQL.FieldByName("NIVEL3").AsInteger = 3) Or (SQL.FieldByName("NIVEL2").AsInteger = 3) Or (SQL.FieldByName("NIVEL1").AsInteger = 3) Then
		If Not CurrentQuery.FieldByName("TABELAPRECO").IsNull Then
			Set INTERFACE = CreateBennerObject("BSPRE001.Rotinas")

			If CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull Then
				INTERFACE.ValoresFxEventos(CurrentSystem, 50, CurrentQuery.FieldByName("CORPOCLINICO").Value, -1, CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value, CurrentQuery.FieldByName("ESTRUTURAFINAL").Value, -1, CurrentQuery.FieldByName("CONVENIO").AsInteger, 0)
			Else
				INTERFACE.ValoresFxEventos(CurrentSystem, 50, CurrentQuery.FieldByName("CORPOCLINICO").Value, -1, CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value, CurrentQuery.FieldByName("ESTRUTURAFINAL").Value, CurrentQuery.FieldByName("REGIMEATENDIMENTO").Value, CurrentQuery.FieldByName("CONVENIO").AsInteger, 0)
			End If
		Else
			bsShowMessage("Para visualizar os preços dos eventos desta configuração de faixa, a tabela genérica deve ser informada !", "I")
		End If
	Else
		bsShowMessage("O 'corpo-clínico' não está configurado na busca de preço !", "I")
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE HANDLE = :HANDLE")

	SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PRESTADORDAENTID")
	SQL.Active = True

	If SQL.FieldByName("PRECO").AsString <> "P" Or SQL.FieldByName("TABPAGAMENTO").AsString <> "1" Then
		CanContinue = False
		bsShowMessage("A configuração do corpo-clínico não permite inserção nesta tabela!", "E")
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
	TABELAPRECO.SetFocus
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
