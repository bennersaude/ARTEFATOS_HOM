﻿'HASH: 9C251F1AA137B7A13805CED6BA69F8CA
'Macro: SAM_PRECOREDE_DOTAC
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOPRECO_OnClick()
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
	Dim SQL As Object
	Dim Nivel As Integer
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active = True

	Nivel = -1

	If CurrentQuery.FieldByName("EVENTO").IsNull Then
		LABELPRECO.Text = ""
	Else
		Set Interface = CreateBennerObject("BSPRE001.Rotinas")

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
				ValorEvento = Interface.ValorEvento(CurrentSystem, vDataBaseChecagemVigencia, 99, -1, CurrentQuery.FieldByName("REDERESTRITA").Value, -1, -1, -1, -1, -1, -1, CurrentQuery.FieldByName("EVENTO").Value, -1, Nivel, CurrentQuery.FieldByName("CONVENIO").AsInteger, "", CurrentQuery.FieldByName("CBOS").AsString)

				LABELPRECO.Text = "Valor do evento nesta vigência: R$ " + Format(ValorEvento, "#,##0.0000")+" ("+Format(ValorEvento,"#,##0.00")+")"
			Else
				LABELPRECO.Text = "Na configuração de busca de preço, não foi definido um nível para a Rede Restrita !"
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
				ValorEvento = Interface.ValorEvento(CurrentSystem, vDataBaseChecagemVigencia, 99, CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").Value, CurrentQuery.FieldByName("REDERESTRITA").Value, -1, -1, -1, -1, -1, -1, CurrentQuery.FieldByName("EVENTO").Value, -1, Nivel, CurrentQuery.FieldByName("CONVENIO").AsInteger, "", CurrentQuery.FieldByName("CBOS").AsString)

				LABELPRECO.Text = "Valor do evento nesta vigência: R$ " + Format(ValorEvento, "#,##0.0000")+" ("+Format(ValorEvento,"#,##0.00")+")"
			Else
				LABELPRECO.Text = "Na configuração de busca de preço, não foi definido um nível para o Prestador da Rede Restrita !"
			End If
		End If
	End If
End Sub

Public Sub BOTAOQUANTIDADES_OnClick()
	Dim QueryRetorno As Object
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vData As String

	If CurrentQuery.State <>3 Then
		Exit Sub
	End If

	If CurrentQuery.FieldByName("EVENTO").IsNull Then
		bsShowMessage("Digite um evento !", "I")
		Exit Sub
	End If

	Set Interface = CreateBennerObject("Procura.Procurar")

	vData = SQLDate( ServerDate)
	vColunas = "SAM_PRECOGENERICO.DESCRICAO|SAM_TGE.DESCRICAO|SAM_PRECOGENERICO_DOTAC.QTDUSHONORARIO|SAM_PRECOGENERICO_DOTAC.QTDUSCUSTOOPERACIONAL|SAM_PRECOGENERICO_DOTAC.FATORFILME|SAM_PRECOGENERICO_DOTAC.PORTEANESTESICO|SAM_PRECOGENERICO_DOTAC.PORTESALA"
	vCriterio = "((SAM_PRECOGENERICO_DOTAC.DATAINICIAL <= " + vData + ") AND (SAM_PRECOGENERICO_DOTAC.DATAFINAL >= " + vData + " OR SAM_PRECOGENERICO_DOTAC.DATAFINAL IS NULL))"
	vCriterio = vCriterio + " AND SAM_PRECOGENERICO_DOTAC.EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	vCampos = "Tabela Genérica|Evento|Qtde US Honorário|Qtde US Custo Operacional|Fator de filme|Porte Anestésico|Porte de Sala"
	vHandle = Interface.Exec(CurrentSystem, "SAM_PRECOGENERICO_DOTAC|SAM_PRECOGENERICO[SAM_PRECOGENERICO_DOTAC.TABELAPRECO = SAM_PRECOGENERICO.HANDLE]|SAM_TGE[SAM_PRECOGENERICO_DOTAC.EVENTO=SAM_TGE.HANDLE]", vColunas, 1, vCampos, vCriterio, "Quantidades", True, "", "")

	Set QueryRetorno = NewQuery

	QueryRetorno.Add("SELECT * FROM SAM_PRECOGENERICO_DOTAC WHERE HANDLE =:HANDLE")
	QueryRetorno.ParamByName("HANDLE").Value = vHandle
	QueryRetorno.Active = True

	CurrentQuery.FieldByName("QTDUSHONORARIO").Value = QueryRetorno.FieldByName("QTDUSHONORARIO").Value
	CurrentQuery.FieldByName("QTDUSCUSTOOPERACIONAL").Value = QueryRetorno.FieldByName("QTDUSCUSTOOPERACIONAL").Value
	CurrentQuery.FieldByName("FATORFILME").Value = QueryRetorno.FieldByName("FATORFILME").Value
	CurrentQuery.FieldByName("PORTEANESTESICO").Value = QueryRetorno.FieldByName("PORTEANESTESICO").Value
	CurrentQuery.FieldByName("PORTESALA").Value = QueryRetorno.FieldByName("PORTESALA").Value

	QueryRetorno.Active = False

	Set Interface = Nothing
End Sub

Public Sub BOTAOREPLICARPORCBOS_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  If (CurrentQuery.State <> 1)  Then
	  bsShowMessage("O registro não pode estar em edição. Confirme ou cancela as alterações!", "I")
	  Exit Sub
  End If

  SessionVar("TabelaDeDotacao") = "SAM_PRECOREDE_DOTAC"
  SessionVar("HandleDotac") = CurrentQuery.FieldByName("HANDLE").AsString

  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
								   1, _
								   "TV_CBOS", _
								   "Replicar por CBO-S", _
								   0, _
								   480, _
								   640, _
								   False, _
								   vsMensagem, _
								   vcContainer)
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim vData As String
	Dim Interface As Object
	Dim SQL As Object
	Dim vColunas, vCriterio, vCampos, vTabela As String

	ShowPopup = False

	If Not CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
		Set SQL = NewQuery

		SQL.Add("SELECT PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE HANDLE = :HANDLE")

		SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsInteger
		SQL.Active = True

		Set Interface = CreateBennerObject("Procura.Procurar")

		vData = SQLDate( ServerDate)
		vColunas = " SAM_TGE.ESTRUTURA|SAM_TGE.DESCRICAO"
		vCriterio = " SAM_TGE.HANDLE  IN ( SELECT DISTINCT GE.EVENTO"
		vCriterio = vCriterio + " FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE  "
		vCriterio = vCriterio + " JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO)  "
		vCriterio = vCriterio + " JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE = EG.ESPECIALIDADE)  "
		vCriterio = vCriterio + " JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE = E.HANDLE)  "
		vCriterio = vCriterio + " LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)  "
		vCriterio = vCriterio + " WHERE PE.DATAINICIAL <= " + vData
		vCriterio = vCriterio + " AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >=" + vData + ")  "
		vCriterio = vCriterio + " AND PE.PRESTADOR = " + SQL.FieldByName("PRESTADOR").AsString
		vCriterio = vCriterio + " AND GE.EVENTO NOT IN (SELECT X.EVENTO  "
		vCriterio = vCriterio + " FROM SAM_PRESTADOR_REGRA X  "
		vCriterio = vCriterio + " WHERE X.REGRAEXCECAO = 'E'  "
		vCriterio = vCriterio + " AND X.PRESTADOR = PE.PRESTADOR  "
		vCriterio = vCriterio + " AND X.DATAFINAL <= " + vData
		vCriterio = vCriterio + " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >=" + vData + "))  "
		vCriterio = vCriterio + " UNION  "
		vCriterio = vCriterio + " SELECT X.EVENTO "
		vCriterio = vCriterio + " FROM SAM_PRESTADOR_REGRA X "
		vCriterio = vCriterio + " WHERE X.REGRAEXCECAO = 'R' "
		vCriterio = vCriterio + " AND X.PRESTADOR = " + SQL.FieldByName("PRESTADOR").AsString
		vCriterio = vCriterio + " AND X.DATAFINAL <= " + vData
		vCriterio = vCriterio + " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >=" + vData + ") "
		vCriterio = vCriterio + " ) "
		vCampos = "Código do evento|Descrição"
		vTabela = "SAM_TGE"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos que o prestador pode executar", True, "")

		SQL.Active = False
	Else
		Set Interface = CreateBennerObject("Procura.Procurar")

		vColunas = " SAM_TGE.ESTRUTURA|SAM_TGE.DESCRICAO"
		vCriterio = " ULTIMONIVEL = 'S'"
		vCampos = "Código do evento|Descrição"
		vTabela = "SAM_TGE"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos que o prestador pode executar", True, "")
	End If

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
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

Public Sub CBOSPESQUISA_OnPopup(ShowPopup As Boolean)
    Dim Interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String
    Dim qCBOS As Object

    ShowPopup = False

    Set Interface = CreateBennerObject("Procura.Procurar")

    vColunas = "TIS_VERSAO.VERSAO|TIS_CBOS.CODIGO|TIS_CBOS.DESCRICAO"
    vCampos = "Versão TISS|Código do CBOS|Descrição do CBOS"
    vHandle = Interface.Exec(CurrentSystem,"TIS_CBOS|TIS_VERSAO[TIS_CBOS.VERSAOTISS = TIS_VERSAO.HANDLE]", vColunas, 2, vCampos, "", "", True, "", CBOSPESQUISA.Text)

    If (vHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("CBOSPESQUISA").Value = vHandle
	End If
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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim vHandle As Long
	Dim vData As String
	Dim SQL As Object
	Dim vPrestador As String
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Dim qCBOS As BPesquisa

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	If CurrentQuery.FieldByName("CBOSPESQUISA").IsNull Then
		CurrentQuery.FieldByName("CBOS").Clear
	Else
		Set qCBOS = NewQuery
		qCBOS.Add("SELECT CODIGO FROM TIS_CBOS WHERE HANDLE = :HANDLE")
		qCBOS.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CBOSPESQUISA").Value
		qCBOS.Active = True
		CurrentQuery.FieldByName("CBOS").Value = qCBOS.FieldByName("CODIGO").Value
		Set qCBOS = Nothing
	End If


	Condicao = "AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	Condicao = Condicao + " AND REDERESTRITA = " + CurrentQuery.FieldByName("REDERESTRITA").AsString

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

	If CurrentQuery.FieldByName("CBOS").IsNull Then
    	Condicao = Condicao + "AND (CBOS IS NULL OR CBOS = '')"
	Else
		Condicao = Condicao + "AND CBOS = '" + CurrentQuery.FieldByName("CBOS").AsString + "'"
	End If

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOREDE_DOTAC", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "REDERESTRITA", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing

	If Not CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
		Set SQL = NewQuery

		SQL.Add("SELECT PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE HANDLE = :HANDLE")

		SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsInteger
		SQL.Active = True

		vPrestador = SQL.FieldByName("PRESTADOR").AsString

		SQL.Active = False

		SQL.Clear

		vData = SQLDate( ServerDate)

		If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
			SQL.Add("SELECT SAM_TGE.HANDLE  ")
			SQL.Add("  FROM SAM_TGE")
			SQL.Add(" WHERE HANDLE IN (SELECT DISTINCT ")
			SQL.Add("                         GE.EVENTO ")
			SQL.Add("                    FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE   ")
			SQL.Add("                    JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO)   ")
			SQL.Add("                    JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE = EG.ESPECIALIDADE)   ")
			SQL.Add("                    JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE = E.HANDLE)   ")
			SQL.Add("               LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)   ")
			SQL.Add("                   WHERE PE.DATAINICIAL <= " + vData)
			SQL.Add("                     AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= " + vData + ") ")
			SQL.Add("                     AND PE.PRESTADOR = " + vPrestador)
			SQL.Add("                     AND GE.EVENTO NOT IN (SELECT X.EVENTO")
			SQL.Add("                                             FROM SAM_PRESTADOR_REGRA X   ")
			SQL.Add("                                            WHERE X.REGRAEXCECAO = 'E'   ")
			SQL.Add("                                              AND X.PRESTADOR = PE.PRESTADOR   ")
			SQL.Add("                                              AND X.DATAINICIAL <=  " + vData)
			SQL.Add("                                              AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vData + ") ")
			SQL.Add("                                          )   ")
			SQL.Add("                  UNION   ")
			SQL.Add("                 SELECT X.EVENTO  ")
			SQL.Add("                   FROM SAM_PRESTADOR_REGRA X  ")
			SQL.Add("                  WHERE X.REGRAEXCECAO = 'R'  ")
			SQL.Add("                    AND X.PRESTADOR = " + vPrestador)
			SQL.Add("                    AND X.DATAINICIAL <=  " + vData)
			SQL.Add("                    AND (X.DATAFINAL IS NULL OR X.DATAFINAL >=  " + vData + ") ")
			SQL.Add("                )  ")
			SQL.Add("   AND SAM_TGE.INATIVO = 'N' ")
			SQL.Add("   AND SAM_TGE.HANDLE  = " + CurrentQuery.FieldByName("EVENTO").AsString)

			SQL.Active = True

			If SQL.EOF Then
				bsShowMessage("Evento não permitido para este prestador.", "E")
				CanContinue = False
				Exit Sub
			End If
		End If
	End If
	If CurrentQuery.FieldByName("CBOSPESQUISA").IsNull Then
		CurrentQuery.FieldByName("CBOS").Value = ""
	End If
End Sub

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
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_AfterScroll()
	LABELPRECO.Text = ""
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
		Case "BOTAOREPLICARPORCBOS"
            SessionVar("TabelaDeDotacao") = "SAM_PRECOREDE_DOTAC"
            SessionVar("HandleDotac") = CurrentQuery.FieldByName("HANDLE").AsString
	End Select
End Sub
