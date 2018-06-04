'HASH: B698674A6A85E0A00B4B1DA0F48547DD
'#Uses "*bsShowMessage"

Dim vCondicao As String

Public Sub TABLE_AfterEdit()
	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere = vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim CONDICAO As String
	Dim vHoraInicial, vHoraFinal As String
	Dim QTab As Object
    Dim EspecificoDll As Object
	'SMS 49152 - Anderson Lonardoni
	'Esta verificação foi tirada do BeforeInsert e colocada no
	'BeforePost para que, no caso de Inserção, já existam valores
	'no CurrentQuery e para funcionar com o Integrator
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
	'SMS 49152 - Fim

	Set QTab = NewQuery

	QTab.Add("SELECT * FROM SAM_TABELAHE_ITEM WHERE TABELAHE = " + CurrentQuery.FieldByName("TABELAHE").AsString)

	QTab.Active = True

	CONDICAO = CONDICAO + "AND (                                                                               "
	CONDICAO = CONDICAO + "   (TABELAHE = " + CurrentQuery.FieldByName("TABELAHE").AsString + ")                   "

	While Not QTab.EOF
		vHoraInicial = SQLDateTime( QTab.FieldByName("HORAINICIAL").AsDateTime)
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

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		CONDICAO = CONDICAO + " AND CONVENIO IS NULL"
	Else
		CONDICAO = CONDICAO + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	'SMS 26799 - Débora Rebello - 22/08/2007 - inicio
	If UserVar("FILTRO_BACEN") <> "" Then
		CONDICAO = CONDICAO + UserVar("FILTRO_BACEN")
	End If
	'SMS 26799 - Débora Rebello - 22/08/2007 - fim

    Set EspecificoDll = CreateBennerObject("ESPECIFICO.UESPECIFICO")
	CONDICAO = CONDICAO + EspecificoDll.CAM_PRO_VerificarVigenciaPrecoPrestadorHorarioEspecial(CurrentSystem, CurrentQuery.TQuery)

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOPRESTADOR_HE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", CONDICAO)

	If Linha = "" Then
		CanContinue = True
	Else
		bsShowMessage(Linha + Chr(13) + "Ou existe horário conflitante para o mesmo dia em outra tabela de HE!", "E")
		CanContinue = False
	End If

	Set Interface = Nothing
	Set EspecificoDll = Nothing
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

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere = vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
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
