'HASH: 556A3AAB5B3F19F58832130AD55E8489
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim CONDICAO As String
	Dim vHoraInicial, vHoraFinal As String
	Dim QTab As Object
	Set QTab = NewQuery

	QTab.Add("SELECT * FROM SAM_TABELAHE_ITEM WHERE TABELAHE = " + CurrentQuery.FieldByName("TABELAHE").AsString)

	QTab.Active = True

	CONDICAO = CONDICAO + "AND (                                                                               "
	CONDICAO = CONDICAO + "   (TABELAHE = " + CurrentQuery.FieldByName("TABELAHE").AsString + ")                   "

	While Not QTab.EOF
		vHoraInicial = SQLDateTime(QTab.FieldByName("HORAINICIAL").AsDateTime)
		vHoraFinal = SQLDateTime(QTab.FieldByName("HORAFINAL").AsDateTime)
		CONDICAO = CONDICAO + "  OR                                                                                "
		CONDICAO = CONDICAO + "   (TABELAHE IN (SELECT A.TABELAHE                                                  "
		CONDICAO = CONDICAO + "                   FROM SAM_TABELAHE_ITEM A                                         "
		CONDICAO = CONDICAO + "                  WHERE ( (HORAINICIAL BETWEEN " + vHoraInicial + " AND " + vHoraFinal + ") "
		CONDICAO = CONDICAO + "                           OR                                                       "
		CONDICAO = CONDICAO + "                          (HORAFINAL   BETWEEN " + vHoraInicial + " AND " + vHoraFinal + ") "
		CONDICAO = CONDICAO + "                           OR                                                       "
		CONDICAO = CONDICAO + "                          (" + vHoraInicial + " BETWEEN A.HORAINICIAL AND A.HORAFINAL)  "
		CONDICAO = CONDICAO + "                           OR                                                       "
		CONDICAO = CONDICAO + "                          (" + vHoraFinal + " BETWEEN A.HORAINICIAL AND A.HORAFINAL)    "
		CONDICAO = CONDICAO + "                        )                                                           "
		CONDICAO = CONDICAO + "                    AND A.TIPODIA = " + QTab.FieldByName("TIPODIA").AsString + "        "
		CONDICAO = CONDICAO + "                )                                                                   "
		CONDICAO = CONDICAO + "   AND TABELAHE <> " + CurrentQuery.FieldByName("TABELAHE").AsString + "                "
		CONDICAO = CONDICAO + "   )                                                                                "

		QTab.Next
	Wend

	CONDICAO = CONDICAO + ")                                                                                   "
	CONDICAO = CONDICAO + " AND REDERESTRITA = " + CurrentQuery.FieldByName("REDERESTRITA").AsString

	If CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
		CONDICAO = CONDICAO + " AND REDERESTRITAPRESTADOR IS NULL "
	Else
		CONDICAO = CONDICAO + " AND REDERESTRITAPRESTADOR = " + CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsString
	End If

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		CONDICAO = CONDICAO + " AND CONVENIO IS NULL"
	Else
		CONDICAO = CONDICAO + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	If (Not CurrentQuery.FieldByName("GRAU").IsNull) Then
    	CONDICAO = CONDICAO + "AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
	Else
		CONDICAO = CONDICAO + "AND GRAU IS NULL"
	End If

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOREDE_HE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "REDERESTRITA", CONDICAO)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha + Chr(13) + "Ou existe horário conflitante para o mesmo dia em outra tabela de HE!", "E")
	End If

	Set Interface = Nothing
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
