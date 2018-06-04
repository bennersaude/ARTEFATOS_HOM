'HASH: 18833CE84F78235CEC8F0C576828F5AA
'MACRO SAM_PRECOPRESTADOR_GRAU
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrau"
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEventoGrau"
'#Uses "*NegociacaoPrecos"

Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False

	If CurrentQuery.FieldByName("GRAU").IsNull Then
		bsShowMessage("Escolha o grau primeiro", "I")
		Exit Sub
	End If

	vHandle = ProcuraEventoGrau(True, EVENTO.Text, CurrentQuery.FieldByName("GRAU").AsInteger)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	Dim vHandleGrau As Long
	Dim INTERFACE As Object
	Dim vColunas, vCampos, vTabela As String
	Set INTERFACE = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = "SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO|SAM_TIPOGRAU.DESCRICAO"
	vCampos = "Código do Grau|Descrição|Tipo do Grau"
	vTabela = "SAM_GRAU|SAM_TIPOGRAU[SAM_TIPOGRAU.HANDLE = SAM_GRAU.TIPOGRAU]"

	'Crys - SMS 76838 - 30/07/2007 - Verifica se o codigo ou a Descrição do Grau é valido, é caso este seja não abri a tela de procura
	If IsNumeric (GRAU.LocateText) Then
		vHandleGrau = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, criaCriterio, "Graus de atuação com configuração de 'preço por grau'", True, GRAU.LocateText)
	Else
		vHandleGrau = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, criaCriterio, "Graus de atuação com configuração de 'preço por grau'", True, GRAU.LocateText)
	End If


	If vHandleGrau <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandleGrau
	End If

	Set INTERFACE = Nothing
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

Public Sub TABLE_AfterEdit()

	Dim vCondicao As String

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

		GRAU.WebLocalWhere = criaCriterio
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vFiltroAdicional As String
  Dim qverifica        As Object
  Dim vAtedias As Integer
  Dim vDeDias As Integer
  Dim vAteAnos As Integer
  Dim vDeAnos As Integer

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

  If VisibleMode Then
    vFiltroAdicional = " AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString

    If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
	  vFiltroAdicional = vFiltroAdicional + " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	Else
	  vFiltroAdicional = vFiltroAdicional + " AND EVENTO IS NULL"
	End If

	'SMS 21799
	If Not CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull Then
	  vFiltroAdicional = vFiltroAdicional + " AND REGIMEATENDIMENTO = " + CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsString
	Else
	  vFiltroAdicional = vFiltroAdicional + " AND REGIMEATENDIMENTO IS NULL "
	End If

	If CurrentQuery.FieldByName("ATEDIAS").IsNull Then
       vAtedias = -1
    Else
       vAtedias = CurrentQuery.FieldByName("ATEDIAS").AsInteger
    End If

    If CurrentQuery.FieldByName("ATEANOS").IsNull Then
       vAteAnos = -1
    Else
       vAteAnos = CurrentQuery.FieldByName("ATEANOS").AsInteger
    End If

    If CurrentQuery.FieldByName("DEDIAS").IsNull Then
       vDeDias = -1
    Else
       vDeDias = CurrentQuery.FieldByName("DEDIAS").AsInteger
    End If

    If CurrentQuery.FieldByName("DEANOS").IsNull Then
       vDeAnos = -1
    Else
       vDeAnos = CurrentQuery.FieldByName("DEANOS").AsInteger
    End If


	CanContinue = ValidacoesBeforePostNegociacaoPreco(CurrentQuery.FieldByName("HANDLE").AsInteger, "SAM_PRECOPRESTADOR_GRAU", "DATAINICIAL", "DATAFINAL", "GRAU", _
	  CurrentQuery.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("EVENTO").AsInteger, CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString, _
	  CurrentQuery.FieldByName("CONVENIO").AsString, vFiltroAdicional, vDeAnos, vDeDias, _
	  vAteAnos, vAtedias, CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger, _
	  CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

    If Not CanContinue Then
      Exit Sub
	End If
  End If


  If Not CurrentQuery.FieldByName("GRAU").IsNull And Not CurrentQuery.FieldByName("EVENTO").IsNull Then
	Dim vHandle As Long
	Dim SQL As Object
	Set SQL = NewQuery

    SQL.Add("SELECT EVENTO         ")
	SQL.Add("  FROM SAM_TGE_GRAU   ")
	SQL.Add(" WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString)
	SQL.Add("   AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString)
	SQL.Active = True

	If SQL.EOF Then
	  bsShowMessage("Este evento não é valido", "E")
	  CanContinue = False
	  Exit Sub
	End If

	Set SQL = Nothing
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

	Dim vCondicao As String

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

		GRAU.WebLocalWhere = criaCriterio
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

Public Function criaCriterio As String
	Dim vsData As String
	Dim vsCriterio As String

	vsData = SQLDate(ServerDate)

	If VisibleMode Then
		vsCriterio = " (SAM_GRAU.PRECOPORGRAU = 'S' OR SAM_GRAU.PRECOPORGRAUDOTACAO = 'S')"
		vsCriterio = vsCriterio + "  AND (( (SAM_GRAU.HANDLE In (SELECT TG.GRAU"
	Else
		vsCriterio = " (A.PRECOPORGRAU = 'S' OR A.PRECOPORGRAUDOTACAO = 'S')"
		vsCriterio = vsCriterio + "  AND (( (A.HANDLE In (SELECT TG.GRAU"
	End If

	vsCriterio = vsCriterio + "                               FROM SAM_TIPOPRESTADOR_GRAU  TG"
	vsCriterio = vsCriterio + "                               JOIN SAM_TIPOPRESTADOR       T  ON (T.HANDLE = TG.TIPOPRESTADOR)"
	vsCriterio = vsCriterio + "                               JOIN SAM_PRESTADOR           P  ON (P.TIPOPRESTADOR = T.HANDLE)"
	vsCriterio = vsCriterio + "                               JOIN SAM_GRAU                G  ON (G.HANDLE = TG.GRAU)"

	If VisibleMode Then
		'vsCriterio = vsCriterio + "                              WHERE P.HANDLE = @PRESTADOR"
		vsCriterio = vsCriterio + "                              WHERE P.HANDLE = " + CurrentQuery.FieldByName("PRESTADOR").AsString
	Else
		vsCriterio = vsCriterio + "                              WHERE P.HANDLE = @CAMPO(PRESTADOR)"
	End If

	vsCriterio = vsCriterio + "                              AND (G.PRECOPORGRAU = 'S' OR G.PRECOPORGRAUDOTACAO = 'S')"
	vsCriterio = vsCriterio + "                             )"
	vsCriterio = vsCriterio + "            AND EXISTS(SELECT TG.GRAU"
	vsCriterio = vsCriterio + "                         FROM SAM_TIPOPRESTADOR_GRAU  TG"
	vsCriterio = vsCriterio + "                         JOIN SAM_TIPOPRESTADOR       T ON (T.HANDLE = TG.TIPOPRESTADOR)"
	vsCriterio = vsCriterio + "                         JOIN SAM_PRESTADOR           P ON (P.TIPOPRESTADOR = T.HANDLE)"

	If VisibleMode Then
		'vsCriterio = vsCriterio + "                        WHERE P.HANDLE = @PRESTADOR"
		vsCriterio = vsCriterio + "                        WHERE P.HANDLE = " + CurrentQuery.FieldByName("PRESTADOR").AsString
	Else
		vsCriterio = vsCriterio + "                        WHERE P.HANDLE = @CAMPO(PRESTADOR)"
	End If

	vsCriterio = vsCriterio + "                      )"
	vsCriterio = vsCriterio + "          )"
	vsCriterio = vsCriterio + "         OR NOT EXISTS(SELECT TG.GRAU"
	vsCriterio = vsCriterio + "                         FROM SAM_TIPOPRESTADOR_GRAU  TG"
	vsCriterio = vsCriterio + "                         JOIN SAM_TIPOPRESTADOR       T ON (T.HANDLE = TG.TIPOPRESTADOR)"
	vsCriterio = vsCriterio + "                         JOIN SAM_PRESTADOR           P ON (P.TIPOPRESTADOR = T.HANDLE)"

	If VisibleMode Then
		'vsCriterio = vsCriterio + "                        WHERE P.HANDLE = @PRESTADOR"
		vsCriterio = vsCriterio + "                        WHERE P.HANDLE = " + CurrentQuery.FieldByName("PRESTADOR").AsString
	Else
		vsCriterio = vsCriterio + "                        WHERE P.HANDLE = @CAMPO(PRESTADOR)"
	End If

	vsCriterio = vsCriterio + "                    )"
	vsCriterio = vsCriterio + "       )"
	vsCriterio = vsCriterio + "  )                                                                                   "

	If VisibleMode Then
		vsCriterio = vsCriterio + " AND SAM_GRAU.HANDLE NOT IN (SELECT PG.GRAU                                           "
	Else
		vsCriterio = vsCriterio + " AND A.HANDLE NOT IN (SELECT PG.GRAU                                           "
	End If
	vsCriterio = vsCriterio + "                           FROM SAM_PRESTADOR_GRAU  PG                                "
	vsCriterio = vsCriterio + "                           JOIN SAM_PRESTADOR           P ON (P.HANDLE = PG.PRESTADOR)"

	If VisibleMode Then
		'vsCriterio = vsCriterio + "                          WHERE P.HANDLE = @PRESTADOR"
		vsCriterio = vsCriterio + "                          WHERE P.HANDLE = " + CurrentQuery.FieldByName("PRESTADOR").AsString
	Else
		vsCriterio = vsCriterio + "                          WHERE P.HANDLE = @CAMPO(PRESTADOR)"
	End If

	vsCriterio = vsCriterio + "                            AND PG.REGRAEXCECAO = 'E'                                 "
	vsCriterio = vsCriterio + "                            AND PG.DATAINICIAL <= " + vsData + "                           "
	vsCriterio = vsCriterio + "                            AND (PG.DATAFINAL IS NULL OR PG.DATAFINAL >= " + vsData + ")   "
	vsCriterio = vsCriterio + "                        )                                                             "

	criaCriterio = vsCriterio
End Function
