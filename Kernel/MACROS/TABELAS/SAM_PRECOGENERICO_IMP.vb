'HASH: 6F2CD0B362695055BEE3820E50448316
'#Uses "*bsShowMessage"

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("TABELAPRECOORIGEM").IsNull Then
		ShowPopup = False
		bsShowMessage("É necessário informar a Tabela Origem", "I")
		Exit Sub
	End If

	Dim Interface As Object
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Set Interface = CreateBennerObject("Procura.Procurar")

	vColunas = "ESTRUTURA|Z_DESCRICAO"
	vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
	vCriterio = vCriterio + " AND SAM_TGE.HANDLE IN (SELECT A.EVENTO FROM SAM_PRECOGENERICO_DOTAC A"
	vCriterio = vCriterio + "                WHERE A.TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECOORIGEM").AsString + ")"
	vCampos = "Evento|Descrição"
	ProcuraEvento = Interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, EVENTOFINAL.Text)

	CurrentQuery.FieldByName("EVENTOFINAL").AsInteger = ProcuraEvento

	Set Interface = Nothing

	ShowPopup = False
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("TABELAPRECOORIGEM").IsNull Then
		ShowPopup = False
		bsShowMessage("É necessário informar a Tabela Origem", "I")
		Exit Sub
	End If

	Dim Interface As Object
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Set Interface = CreateBennerObject("Procura.Procurar")

	vColunas = "ESTRUTURA|Z_DESCRICAO"
	vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
	vCriterio = vCriterio + " AND SAM_TGE.HANDLE IN (SELECT A.EVENTO FROM SAM_PRECOGENERICO_DOTAC A"
	vCriterio = vCriterio + "                WHERE A.TABELAPRECO = " + CurrentQuery.FieldByName("TABELAPRECOORIGEM").AsString + ")"
	vCampos = "Evento|Descrição"
	ProcuraEvento = Interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, EVENTOINICIAL.Text)

	CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = ProcuraEvento

	Set Interface = Nothing

	ShowPopup = False
End Sub

Public Sub TABLE_AfterScroll()
  EVENTOINICIAL.WebLocalWhere = " A.HANDLE IN (SELECT X.EVENTO FROM SAM_PRECOGENERICO_DOTAC X" + _
	                            "              WHERE X.TABELAPRECO = @CAMPO(TABELAPRECOORIGEM))"
  EVENTOFINAL.WebLocalWhere   = " A.HANDLE IN (SELECT X.EVENTO FROM SAM_PRECOGENERICO_DOTAC X" + _
	                            "              WHERE X.TABELAPRECO = @CAMPO(TABELAPRECOORIGEM))"
End Sub

Public Sub TABLE_BeforeCancel(CanContinue As Boolean)

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If CurrentQuery.FieldByName("PROCESSADO").AsString = "S" Then
		CanContinue = False
		bsShowMessage("O parâmentro de importação já foi processado. Exclusão não permitida", "E")
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If CurrentQuery.FieldByName("PROCESSADO").AsString = "S" Then
		CanContinue = False
		bsShowMessage("O parâmentro de importação já foi processado. Alteração não permitida", "E")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If Not (CurrentQuery.FieldByName("DATAFINAL").IsNull) And _
		   (CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) Then
		CanContinue = False
		bsShowMessage("A Data Final não pode ser inferior à Data Inicial", "E")
		Exit Sub
	End If

	Dim SQL As Object
	Dim vTabelaUS As String
	Dim vTabelaFilme As String
	Set SQL = NewQuery

	SQL.Clear

	SQL.Add("SELECT TABELAUS, TABELAFILME")
	SQL.Add("FROM SAM_PRECOGENERICO")
	SQL.Add("WHERE HANDLE = :HPRECOGENERICO")

	SQL.ParamByName("HPRECOGENERICO").Value = RecordHandleOfTable("SAM_PRECOGENERICO")
	SQL.Active = True

	vTabelaUS = SQL.FieldByName("TABELAUS").AsInteger
	vTabelaFilme = SQL.FieldByName("TABELAFILME").AsInteger

	SQL.Clear

	SQL.Add("SELECT TABELAUS, TABELAFILME")
	SQL.Add("FROM SAM_PRECOGENERICO")
	SQL.Add("WHERE HANDLE = :HPRECOGENERICO")

	SQL.ParamByName("HPRECOGENERICO").Value = CurrentQuery.FieldByName("TABELAPRECOORIGEM").AsInteger
	SQL.Active = True

	If vTabelaUS <> SQL.FieldByName("TABELAUS").AsInteger Then
		CanContinue = False
		bsShowMessage("A tabela de US de origem não pode ser diferente da tabela atual", "E")
		Set SQL = Nothing
		Exit Sub
	End If

	If vTabelaFilme <> SQL.FieldByName("TABELAFILME").AsInteger Then
		CanContinue = False
		bsShowMessage("A tabela de Filme de origem não pode ser diferente da tabela atual", "E")
		Set SQL = Nothing
		Exit Sub
	End If

	Dim vEventoInicial As String
	Dim vEventoFinal As String
	Set SQL = NewQuery

	SQL.Add("SELECT ESTRUTURA")
	SQL.Add("FROM SAM_TGE")
	SQL.Add("WHERE HANDLE = :HEVENTO")

	SQL.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
	SQL.Active = True

	vEventoInicial = SQL.FieldByName("ESTRUTURA").AsString

	SQL.Active = False
	SQL.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
	SQL.Active = True

	vEventoFinal = SQL.FieldByName("ESTRUTURA").AsString

	SQL.Clear

	SQL.Add("SELECT TABELAUS, TABELAFILME")
	SQL.Add("FROM SAM_PRECOGENERICO")
	SQL.Add("WHERE HANDLE = :HPRECOGENERICO")

	If VisibleMode Then
		SQL.ParamByName("HPRECOGENERICO").Value = RecordHandleOfTable("SAM_PRECOGENERICO")
	Else
		SQL.ParamByName("HPRECOGENERICO").Value = CurrentQuery.FieldByName("TABELAPRECO").AsInteger
	End If

	SQL.Active = True

	If vEventoFinal <vEventoInicial Then
		CanContinue = False
		Set SQL = Nothing
		bsShowMessage("O Evento Final não pode ser inferior ao Evento Inicial", "E")
		Exit Sub
	End If

	SQL.Clear

	SQL.Add("SELECT A.HANDLE")
	SQL.Add("FROM SAM_PRECOGENERICO_IMP A, SAM_TGE B1, SAM_TGE B2")
	SQL.Add("WHERE A.HANDLE <> :HTABELAIMP")
	SQL.Add("  AND TABELAPRECO = :HTABELAPRECO")
	SQL.Add("  AND ( (A.DATAINICIAL <= :DATAINICIAL AND A.DATAFINAL IS NULL) OR")
	SQL.Add("        (A.DATAINICIAL <= :DATAINICIAL AND A.DATAFINAL >= :DATAINICIAL) OR")
	SQL.Add("        (:DATAINICIAL <= A.DATAINICIAL AND :DATAFINAL >= A.DATAINICIAL) ) ")
	SQL.Add("  AND B1.HANDLE = A.EVENTOINICIAL")
	SQL.Add("  AND B2.HANDLE = A.EVENTOFINAL")
	'fernando sms 54093
	SQL.Add("  AND A.PROCESSADO <> 'S' ")
	SQL.Add("  AND ((B1.ESTRUTURA <= :EVENTOINICIAL AND B2.ESTRUTURA >= :EVENTOINICIAL) OR")
	SQL.Add("       (:EVENTOINICIAL <= B1.ESTRUTURA AND :EVENTOFINAL >= B1.ESTRUTURA))")

	SQL.ParamByName("HTABELAIMP").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.ParamByName("HTABELAPRECO").Value = CurrentQuery.FieldByName("TABELAPRECO").AsInteger
	SQL.ParamByName("DATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
	SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
	SQL.ParamByName("EVENTOINICIAL").Value = vEventoInicial
	SQL.ParamByName("EVENTOFINAL").Value = vEventoFinal
	SQL.Active = True

	If Not SQL.EOF Then
		CanContinue = False
		Set SQL = Nothing
		bsShowMessage("Já existem eventos deste intervalo", "E")
	End If
End Sub
