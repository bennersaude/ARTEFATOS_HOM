'HASH: D6B11ACEA2611D5F6B45509214CFD65A
'Macro: SAM_PRECOREDE_FX
'#Uses "*ProcuraTabelaFilme"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaGenerica"
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit


Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
	Dim INTERFACE  As Object
	Dim vsData     As String
	Dim vsColunas  As String
	Dim vsCriterio As String
	Dim vsCampos   As String
	Dim vsTabela   As String
	Dim viHandle   As Long

	ShowPopup = False

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

	vsCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString

	If (Not CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull) Then
		Dim qPrestador As Object
		Set qPrestador = NewQuery

		qPrestador.Add("SELECT A.PRESTADOR,                                            ")
		qPrestador.Add("       B.ASSOCIACAO                                            ")
		qPrestador.Add("  FROM SAM_REDERESTRITA_PRESTADOR A                            ")
		qPrestador.Add("  JOIN SAM_PRESTADOR              B ON (B.HANDLE = A.PRESTADOR)")
		qPrestador.Add(" WHERE A.HANDLE = :HANDLE                                      ")

		qPrestador.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsInteger
		qPrestador.Active = True

		vsData = SQLDate(ServerDate)

		vsCriterio = vsCriterio + " AND "

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
	End If

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vsCriterio = vsCriterio + " AND (SAM_TGE.HANDLE IN (SELECT EVENTO
		vsCriterio = vsCriterio + "                           FROM SAM_PRECOGENERICO_DOTAC
		vsCriterio = vsCriterio + "                          WHERE TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString + ")) "
	End If

	vsTabela  = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vsCampos  = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vsColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	Set INTERFACE = CreateBennerObject("Procura.Procurar")

	viHandle = INTERFACE.Exec(CurrentSystem, vsTabela, vsColunas, 1, vsCampos, vsCriterio, "Tabela Geral de Eventos", True, EVENTOINICIAL.Text)

	If (viHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = viHandle
	End If

	Set INTERFACE = Nothing
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
	Dim INTERFACE  As Object
	Dim vsData     As String
	Dim vsColunas  As String
	Dim vsCriterio As String
	Dim vsCampos   As String
	Dim vsTabela   As String
	Dim viHandle   As Long

	ShowPopup = False

    If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
    End If

    If CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = 0 Then
 	  bsShowMessage("Necessário escolher Evento Inicial antes de selecionar Evento Final", "E")
	  Exit Sub
    End If

	vsCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString

	If (Not CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull) Then
		Dim qPrestador As Object
		Set qPrestador = NewQuery

		qPrestador.Add("SELECT A.PRESTADOR,                                            ")
		qPrestador.Add("       B.ASSOCIACAO                                            ")
		qPrestador.Add("  FROM SAM_REDERESTRITA_PRESTADOR A                            ")
		qPrestador.Add("  JOIN SAM_PRESTADOR              B ON (B.HANDLE = A.PRESTADOR)")
		qPrestador.Add(" WHERE A.HANDLE = :HANDLE                                      ")

		qPrestador.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsInteger
		qPrestador.Active = True

		vsData = SQLDate(ServerDate)

		vsCriterio = vsCriterio + " AND "

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
	End If

	If (Not CurrentQuery.FieldByName("TABELAPRECO").IsNull) Then
		vsCriterio = vsCriterio + " AND (SAM_TGE.HANDLE IN (SELECT EVENTO
		vsCriterio = vsCriterio + "                           FROM SAM_PRECOGENERICO_DOTAC
		vsCriterio = vsCriterio + "                          WHERE TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECO").AsString + ")) "
	End If

	vsTabela  = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vsCampos  = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vsColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	Set INTERFACE = CreateBennerObject("Procura.Procurar")

	viHandle = INTERFACE.Exec(CurrentSystem, vsTabela, vsColunas, 1, vsCampos, vsCriterio, "Tabela Geral de Eventos", True, EVENTOFINAL.Text)

	If (viHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTOFINAL").AsInteger = viHandle
	End If

	Set INTERFACE = Nothing
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

Public Sub TABELAFILME_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaFilme(TABELAFILME.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAFILME").Value = vHandle
	End If
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaUS(TABELAUS.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAUS").Value = vHandle
	End If
End Sub

Public Sub TABELAPRECO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaGenerica(TABELAPRECO.Text)

	If vHandle <>0 Then
		CurrentQuery.FieldByName("TABELAPRECO").Value = vHandle
		CurrentQuery.FieldByName("EVENTOINICIAL").Value = Null
		CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = Null
		CurrentQuery.FieldByName("EVENTOFINAL").Value = Null
		CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Null
	End If
End Sub

Public Sub TABLE_AfterScroll()

	If WebMode Then
	  EVENTOINICIAL.WebLocalWhere = " A.MASCARATGE = @CAMPO(MASCARATGE)"
	  EVENTOFINAL.WebLocalWhere = " A.MASCARATGE = @CAMPO(MASCARATGE)"
	End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim INTERFACE As Object
	Dim Linha As String
	Dim Condicao As String

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
	Dim Estrutura, EstruturaI, EstruturaF As String
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

	' Checar Vigencia
	Condicao = " REDERESTRITA = " + CurrentQuery.FieldByName("REDERESTRITA").AsString

	If CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
		Condicao = Condicao + " AND REDERESTRITAPRESTADOR IS NULL "
	Else
		Condicao = Condicao + " AND REDERESTRITAPRESTADOR = " + CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsString
	End If

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

	Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = INTERFACE.EventoFx(CurrentSystem, "SAM_PRECOREDE_FX", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set INTERFACE = Nothing

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

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		CanContinue = False
		bsShowMessage("Registro finalizado não pode ser alterado", "E")
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub BOTAOVALORES_OnClick()
	Dim INTERFACE As Object
	Dim SQL As Object
	Dim Nivel As Integer
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active = True

	Nivel = -1

	If Not CurrentQuery.FieldByName("TABELAPRECO").IsNull Then
		Set INTERFACE = CreateBennerObject("BSPRE001.Rotinas")

		If CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
			If SQL.FieldByName("NIVEL1").AsInteger = 2 Then
				Nivel = 1
			ElseIf SQL.FieldByName("NIVEL2").AsInteger = 2 Then
				Nivel = 2
			ElseIf SQL.FieldByName("NIVEL3").AsInteger = 2 Then
				Nivel = 3
			ElseIf SQL.FieldByName("NIVEL4").AsInteger = 2 Then
				Nivel = 4
			ElseIf SQL.FieldByName("NIVEL5").AsInteger = 2 Then
				Nivel = 5
			ElseIf SQL.FieldByName("NIVEL6").AsInteger = 2 Then
				Nivel = 6
			ElseIf SQL.FieldByName("NIVEL7").AsInteger = 2 Then
				Nivel = 7
			ElseIf SQL.FieldByName("NIVEL8").AsInteger = 2 Then
				Nivel = 8
			End If

			If Nivel <> -1 Then
				INTERFACE.ValoresFxEventos(CurrentSystem, 60, CurrentQuery.FieldByName("REDERESTRITA").Value, -1, CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value, CurrentQuery.FieldByName("ESTRUTURAFINAL").Value, -1, CurrentQuery.FieldByName("CONVENIO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("MASCARATGE").AsInteger)
			Else
				bsShowMessage("Na configuração da busca do preço, não foi definido um nível para a Rede Restrita !", "I")
			End If
		Else
			If SQL.FieldByName("NIVEL1").AsInteger = 1 Then
				Nivel = 1
			ElseIf SQL.FieldByName("NIVEL2").AsInteger = 1 Then
				Nivel = 2
			ElseIf SQL.FieldByName("NIVEL3").AsInteger = 1 Then
				Nivel = 3
			ElseIf SQL.FieldByName("NIVEL4").AsInteger = 1 Then
				Nivel = 4
			ElseIf SQL.FieldByName("NIVEL5").AsInteger = 1 Then
				Nivel = 5
			ElseIf SQL.FieldByName("NIVEL6").AsInteger = 1 Then
				Nivel = 6
			ElseIf SQL.FieldByName("NIVEL7").AsInteger = 1 Then
				Nivel = 7
			ElseIf SQL.FieldByName("NIVEL8").AsInteger = 1 Then
				Nivel = 8
			End If

			If Nivel <> -1 Then
				INTERFACE.ValoresFxEventos(CurrentSystem, 70, CurrentQuery.FieldByName("REDERESTRITA").Value, CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").Value, CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value, CurrentQuery.FieldByName("ESTRUTURAFINAL").Value, -1, CurrentQuery.FieldByName("CONVENIO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("MASCARATGE").AsInteger)
			Else
				bsShowMessage("Na configuração da busca do preço, não foi definido um nível para o Prestador na Rede Restrita !", "I")
			End If
		End If
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
