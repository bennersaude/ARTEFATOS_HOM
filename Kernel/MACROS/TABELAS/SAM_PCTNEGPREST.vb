'HASH: 6C63DBC621E9BDEB9642352AE54502B0
'Macro: SAM_PCTNEGPREST
'Mauricio Ibelli -04/05/2001 -sms 2226 -Selecionar grau validos
'#Uses "*bsShowMessage"
'#Uses "*NegociacaoPrecos"

Public Sub BOTAOGERARRELATORIO_OnClick()
If CurrentQuery.State <> 1 Then
	    bsShowMessage("O registro está em edição.","I")
    Else

    	Dim RelatorioHandle As Long
		Dim QueryBuscaHandleRelatorio As Object


		Set QueryBuscaHandleRelatorio=NewQuery

		QueryBuscaHandleRelatorio.Add("SELECT RELATORIOPACOTE FROM SAM_PARAMETROSPRESTADOR")
    	        QueryBuscaHandleRelatorio.Active=False
   		QueryBuscaHandleRelatorio.Active=True
   		RelatorioHandle=QueryBuscaHandleRelatorio.FieldByName("RELATORIOPACOTE").AsInteger

		If (RelatorioHandle = 0) Then
		 bsShowMessage("Relatório não está parametrizado","I")
		 CanContinue = False
		Else
		 ReportPreview(RelatorioHandle,"", False, False)
		End If

	    Set QueryBuscaHandleRelatorio=Nothing
	End If
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vData As String
  Dim Interface As Object
  Dim vColunas, vCriterio, vCampos, vTabela As String
  Dim qPrestador As Object

  ShowPopup = False



  Set Interface =CreateBennerObject("Procura.Procurar")
  vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
  vColunas ="SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"
  vCriterio = criarCriterio
  vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
  vHandle =Interface.Exec(CurrentSystem,vTabela ,vColunas, 1, vCampos, vCriterio, "Eventos que que o prestador pode executar",True,"","")


'  vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
'  vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
'  vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos que que o prestador pode executar", True, "")

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
    CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  End If

End Sub

Public Sub GRAUAGERAR_OnPopup(ShowPopup As Boolean)
  If VisibleMode Then
  	EVENTO.LocalWhere = vCriterio
  	If CurrentQuery.FieldByName("EVENTO").AsString <> "" Then
    	GRAUAGERAR.LocalWhere = "ORIGEMVALOR ='7' AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
    Else
         bsShowMessage("O Campo evento deve ser preenchido.", "I")
         ShowPopup = False
    End If
  End If

End Sub

Public Sub TABLE_AfterEdit()
  Dim qPrestador As Object
  Set qPrestador = NewQuery

  qPrestador.Active = False
  qPrestador.Add("SELECT ASSOCIACAO FROM SAM_PRESTADOR WHERE HANDLE=:PRESTADOR")
  qPrestador.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  qPrestador.Active = True

  If qPrestador.FieldByName("ASSOCIACAO").AsString <> "S" Then
    vCriterio = vCriterio + " SAM_TGE.HANDLE  IN ( SELECT DISTINCT GE.EVENTO"
    vCriterio = vCriterio + " FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE  "
    vCriterio = vCriterio + " JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO)  "
    vCriterio = vCriterio + " JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE = EG.ESPECIALIDADE)  "
    vCriterio = vCriterio + " JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE = E.HANDLE)  "
    vCriterio = vCriterio + " LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)  "
    vCriterio = vCriterio + " WHERE PE.DATAINICIAL <= " +SQLDate(ServerDate)
    vCriterio = vCriterio + " AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >=" + SQLDate(ServerDate) + ")  "
    vCriterio = vCriterio + " AND PE.PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
    vCriterio = vCriterio + " AND GE.EVENTO NOT IN (SELECT X.EVENTO  "
    vCriterio = vCriterio + " FROM SAM_PRESTADOR_REGRA X  "
    vCriterio = vCriterio + " WHERE X.REGRAEXCECAO = 'E'  "
    vCriterio = vCriterio + " AND X.PRESTADOR = PE.PRESTADOR  "
    vCriterio = vCriterio + " AND X.DATAINICIAL <= " + SQLDate(ServerDate)
    vCriterio = vCriterio + " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >=" + SQLDate(ServerDate) + "))  "
    vCriterio = vCriterio + " UNION  "
    vCriterio = vCriterio + " SELECT X.EVENTO "
    vCriterio = vCriterio + " FROM SAM_PRESTADOR_REGRA X "
    vCriterio = vCriterio + " WHERE X.REGRAEXCECAO = 'R' "
    vCriterio = vCriterio + " AND X.PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
    vCriterio = vCriterio + " AND X.DATAINICIAL <= " + SQLDate(ServerDate)
    vCriterio = vCriterio + " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >=" + SQLDate(ServerDate) + ") "
    vCriterio = vCriterio + " ) "
  Else
    vCriterio = "SAM_TGE.ULTIMONIVEL='S'"
  End If

  If VisibleMode Then
  	EVENTO.LocalWhere = vCriterio
  Else
  	EVENTO.WebLocalWhere = vCriterio
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim vAtedias As Integer
  Dim vDeDias As Integer
  Dim vAteAnos As Integer
  Dim vDeAnos As Integer

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Condicao = " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
  Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
  If Not CurrentQuery.FieldByName("GRAUAGERAR").IsNull Then
	Condicao = Condicao + " AND GRAUAGERAR = " + CurrentQuery.FieldByName("GRAUAGERAR").AsString
  End If

  If (CurrentQuery.FieldByName("ACOMODACAO").IsNull) Then
	Condicao = Condicao + " AND ACOMODACAO IS NULL "
  Else
    Condicao = Condicao + " AND ACOMODACAO = " + CurrentQuery.FieldByName("ACOMODACAO").AsString
  End If

  Linha = Interface.Vigencia(CurrentSystem, "SAM_PCTNEGPREST", _
          "DATAINICIAL", _
          "DATAFINAL", _
          CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
          CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
          "PRESTADOR", _
          Condicao)

  If Linha <> "" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

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

  CanContinue = ValidacoesBeforePostNegociacaoPreco(CurrentQuery.FieldByName("HANDLE").AsInteger, "SAM_PCTNEGPREST", "DATAINICIAL", "DATAFINAL", "PRESTADOR", _
	CurrentQuery.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("EVENTO").AsInteger, "", _
	CurrentQuery.FieldByName("CONVENIO").AsString, Condicao, vDeAnos, vDeDias, _
	vAteAnos, vAtedias, CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger, _
	CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

    If Not CanContinue Then
      Exit Sub
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
	Dim qPrestador As BPesquisa

	Set qPrestador = NewQuery

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")

		CanContinue = False

		Exit Sub
	End If

	Dim vCondicao As String
	Dim vCriterio As String

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

	qPrestador.Active = False

	qPrestador.Add("SELECT ASSOCIACAO")
	qPrestador.Add("  FROM SAM_PRESTADOR")
	qPrestador.Add(" WHERE HANDLE=:PRESTADOR")

	qPrestador.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

	qPrestador.Active = True

	If (qPrestador.FieldByName("ASSOCIACAO").AsString <> "S") Then
		vCriterio = vCriterio + "A.HANDLE"
		vCriterio = vCriterio + " IN (  SELECT DISTINCT GE.EVENTO																	 "
		vCriterio = vCriterio + "		  FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE  													 "
		vCriterio = vCriterio + "		  JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE             = GE.ESPECIALIDADEGRUPO)  "
		vCriterio = vCriterio + "		  JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE              = EG.ESPECIALIDADE)  	 "
		vCriterio = vCriterio + "		  JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE      = E.HANDLE)  			 "
		vCriterio = vCriterio + "	 LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)  			 "
		vCriterio = vCriterio + "		 WHERE PE.DATAINICIAL <= " + SQLDate(ServerDate)
		vCriterio = vCriterio + "		   AND (PE.DATAFINAL IS NULL																 "
		vCriterio = vCriterio + "			OR  PE.DATAFINAL  >= " + SQLDate(ServerDate) + ")"
		vCriterio = vCriterio + "		   AND PE.PRESTADOR	   = @CAMPO(PRESTADOR)													 "
		vCriterio = vCriterio + "		   AND GE.EVENTO NOT IN (SELECT X.EVENTO  													 "
		vCriterio = vCriterio + "								   FROM SAM_PRESTADOR_REGRA X  										 "
		vCriterio = vCriterio + "								  WHERE X.REGRAEXCECAO = 'E'  										 "
		vCriterio = vCriterio + "									AND X.PRESTADOR    = PE.PRESTADOR  								 "
		vCriterio = vCriterio + "									AND X.DATAINICIAL <= " + SQLDate(ServerDate)
		vCriterio = vCriterio + "									AND (X.DATAFINAL IS NULL"
		vCriterio = vCriterio + "									 OR  X.DATAFINAL  >= " + SQLDate(ServerDate) + "))"
		vCriterio = vCriterio + "								  UNION  															 "
		vCriterio = vCriterio + "								 SELECT X.EVENTO 													 "
		vCriterio = vCriterio + "								   FROM SAM_PRESTADOR_REGRA X 										 "
		vCriterio = vCriterio + "								  WHERE X.REGRAEXCECAO = 'R' 										 "
		vCriterio = vCriterio + "									AND X.PRESTADOR	   = @CAMPO(PRESTADOR)							 "
		vCriterio = vCriterio + "									AND X.DATAINICIAL <= " + SQLDate(ServerDate)
		vCriterio = vCriterio + "									AND (X.DATAFINAL IS NULL"
		vCriterio = vCriterio + "									 OR  X.DATAFINAL  >= " + SQLDate(ServerDate) + ")"
		vCriterio = vCriterio + ")"
	Else
		vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
	End If

	If WebMode Then EVENTO.WebLocalWhere = vCriterio
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String
	Dim qPrestador As BPesquisa

	Set qPrestador = NewQuery

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")

		CanContinue = False

		Exit Sub
	End If

	Dim vCondicao As String
	Dim vCriterio As String

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

	qPrestador.Active = False

	qPrestador.Add("SELECT ASSOCIACAO")
	qPrestador.Add("  FROM SAM_PRESTADOR")
	qPrestador.Add(" WHERE HANDLE=:PRESTADOR")

	qPrestador.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

	qPrestador.Active = True

	If (qPrestador.FieldByName("ASSOCIACAO").AsString <> "S") Then
		vCriterio = vCriterio + "A.HANDLE"
		vCriterio = vCriterio + " IN (   SELECT DISTINCT GE.EVENTO																	 "
		vCriterio = vCriterio + "		  FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE  													 "
		vCriterio = vCriterio + "		  JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE             = GE.ESPECIALIDADEGRUPO)  "
		vCriterio = vCriterio + "		  JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE              = EG.ESPECIALIDADE)  	 "
		vCriterio = vCriterio + "		  JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE      = E.HANDLE)  			 "
		vCriterio = vCriterio + "	 LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)  			 "
		vCriterio = vCriterio + "		 WHERE PE.DATAINICIAL <= " + SQLDate(ServerDate)
		vCriterio = vCriterio + "		   AND (PE.DATAFINAL IS NULL																 "
		vCriterio = vCriterio + "			OR  PE.DATAFINAL  >= " + SQLDate(ServerDate) + ")"
		vCriterio = vCriterio + "		   AND PE.PRESTADOR	   = @CAMPO(PRESTADOR)													 "
		vCriterio = vCriterio + "		   AND GE.EVENTO NOT IN (SELECT X.EVENTO  													 "
		vCriterio = vCriterio + "								   FROM SAM_PRESTADOR_REGRA X  										 "
		vCriterio = vCriterio + "								  WHERE X.REGRAEXCECAO = 'E'  										 "
		vCriterio = vCriterio + "									AND X.PRESTADOR    = PE.PRESTADOR  								 "
		vCriterio = vCriterio + "									AND X.DATAINICIAL <= " + SQLDate(ServerDate)
		vCriterio = vCriterio + "									AND (X.DATAFINAL IS NULL"
		vCriterio = vCriterio + "									 OR  X.DATAFINAL  >= " + SQLDate(ServerDate) + "))"
		vCriterio = vCriterio + "								  UNION  															 "
		vCriterio = vCriterio + "								 SELECT X.EVENTO 													 "
		vCriterio = vCriterio + "								   FROM SAM_PRESTADOR_REGRA X 										 "
		vCriterio = vCriterio + "								  WHERE X.REGRAEXCECAO = 'R' 										 "
		vCriterio = vCriterio + "									AND X.PRESTADOR	   = @CAMPO(PRESTADOR)							 "
		vCriterio = vCriterio + "									AND X.DATAINICIAL <= " + SQLDate(ServerDate)
		vCriterio = vCriterio + "									AND (X.DATAFINAL IS NULL"
		vCriterio = vCriterio + "									 OR  X.DATAFINAL  >= " + SQLDate(ServerDate) + ")"
		vCriterio = vCriterio + ")"
	Else
		vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
	End If

	If WebMode Then EVENTO.WebLocalWhere = vCriterio
End Sub

Public Sub TABLE_AfterScroll()
  '-------------VALOR TOTAL DO PACOTE ------------------------------------------------
  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    Dim Interface As Object
    Dim valorpacte As Currency
    Set Interface = CreateBennerObject("BSPRE001.Rotinas")

    valorpacte = Interface.ValorTotalPacote(CurrentSystem, "SAM_PCTNEGPREST", CurrentQuery.FieldByName("HANDLE").Value)
    VALORPACOTE.Text = "Valor total do pacote: R$ " + Format(valorpacte, "#,##0.00")
  Else
    VALORPACOTE.Text = " "
  End If
  '-------------VALOR TOTAL DO PACOTE ------------------------------------------------

 If VisibleMode Then
  	EVENTO.LocalWhere = vCriterio
  	If CurrentQuery.FieldByName("EVENTO").AsString <> "" Then
    	GRAUAGERAR.LocalWhere = "ORIGEMVALOR ='7' AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
  	End If
  Else
  	If WebMenuCode = "T5674" Then
  		EVENTO.ReadOnly = True
  	End If
  	If WebMenuCode = "T1303" Then
  		PRESTADOR.ReadOnly = True
  	End If
	GRAUAGERAR.WebLocalWhere = "ORIGEMVALOR ='7' AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO))
  	EVENTO.WebLocalWhere = vCriterio
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

  Dim qPrestador As Object
  Set qPrestador = NewQuery

  qPrestador.Active = False
  qPrestador.Add("SELECT ASSOCIACAO FROM SAM_PRESTADOR WHERE HANDLE=:PRESTADOR")
  qPrestador.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  qPrestador.Active = True

  If qPrestador.FieldByName("ASSOCIACAO").AsString <> "S" Then
    vCriterio = vCriterio + " SAM_TGE.HANDLE  IN ( SELECT DISTINCT GE.EVENTO"
    vCriterio = vCriterio + " FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE  "
    vCriterio = vCriterio + " JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO)  "
    vCriterio = vCriterio + " JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE = EG.ESPECIALIDADE)  "
    vCriterio = vCriterio + " JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE = E.HANDLE)  "
    vCriterio = vCriterio + " LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)  "
    vCriterio = vCriterio + " WHERE PE.DATAINICIAL <= " +SQLDate(ServerDate)
    vCriterio = vCriterio + " AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >=" + SQLDate(ServerDate) + ")  "
    vCriterio = vCriterio + " AND PE.PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
    vCriterio = vCriterio + " AND GE.EVENTO NOT IN (SELECT X.EVENTO  "
    vCriterio = vCriterio + " FROM SAM_PRESTADOR_REGRA X  "
    vCriterio = vCriterio + " WHERE X.REGRAEXCECAO = 'E'  "
    vCriterio = vCriterio + " AND X.PRESTADOR = PE.PRESTADOR  "
    vCriterio = vCriterio + " AND X.DATAINICIAL <= " + SQLDate(ServerDate)
    vCriterio = vCriterio + " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >=" + SQLDate(ServerDate) + "))  "
    vCriterio = vCriterio + " UNION  "
    vCriterio = vCriterio + " SELECT X.EVENTO "
    vCriterio = vCriterio + " FROM SAM_PRESTADOR_REGRA X "
    vCriterio = vCriterio + " WHERE X.REGRAEXCECAO = 'R' "
    vCriterio = vCriterio + " AND X.PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
    vCriterio = vCriterio + " AND X.DATAINICIAL <= " + SQLDate(ServerDate)
    vCriterio = vCriterio + " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >=" + SQLDate(ServerDate) + ") "
    vCriterio = vCriterio + " ) "
  Else
    vCriterio = "SAM_TGE.ULTIMONIVEL='S'"
  End If

  If VisibleMode Then
  	EVENTO.LocalWhere = vCriterio
  Else
  	EVENTO.WebLocalWhere = vCriterio
  End If

End Sub

'MILANI -SMS -22609
Public Sub BOTAOINCLUIRITENS_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("BSPRE009.ROTINAS")

  Interface.ItensPacotes(CurrentSystem, "SAM_PCTNEGPREST_GRAU", CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set Interface = Nothing
End Sub


Public Function criarCriterio As String
'	vCriterio = " A.HANDLE IN (SELECT HANDLE FROM SAM_TGE WHERE ULTIMONIVEL = 'S') OR A.HANDLE IN (SELECT HANDLE FROM SAM_CBHPM WHERE ULTIMONIVEL = 'S')"
'	criarCriterio = vCriterio


    Dim qPrestador As Object
    Set qPrestador = NewQuery

    qPrestador.Clear
	qPrestador.Add("SELECT ASSOCIACAO")
	qPrestador.Add("  FROM SAM_PRESTADOR")
	qPrestador.Add(" WHERE HANDLE=:PRESTADOR")

	qPrestador.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

	qPrestador.Active = True

	If (qPrestador.FieldByName("ASSOCIACAO").AsString <> "S") Then
		vCriterio = vCriterio + "SAM_TGE.HANDLE"
		vCriterio = vCriterio + " IN (   SELECT DISTINCT GE.EVENTO																	 										"
		vCriterio = vCriterio + "		  FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE  													 										"
		vCriterio = vCriterio + "		  JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE             = GE.ESPECIALIDADEGRUPO) 										"
		vCriterio = vCriterio + "		  JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE              = EG.ESPECIALIDADE)  	 										"
		vCriterio = vCriterio + "		  JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE      = E.HANDLE)  			 										"
		vCriterio = vCriterio + "	 LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)  												 	"
		vCriterio = vCriterio + "		 WHERE PE.DATAINICIAL <= "+SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)+"                                          	"
		vCriterio = vCriterio + "		   AND (PE.DATAFINAL IS NULL																 										"
		vCriterio = vCriterio + "			OR  PE.DATAFINAL  >= "+SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)+")                                           "
		vCriterio = vCriterio + "		   AND PE.PRESTADOR	   = "+CurrentQuery.FieldByName("PRESTADOR").AsString+"													 		"
		vCriterio = vCriterio + "		   AND GE.EVENTO NOT IN (SELECT X.EVENTO  													 										"
		vCriterio = vCriterio + "								   FROM SAM_PRESTADOR_REGRA X  										 										"
		vCriterio = vCriterio + "								  WHERE X.REGRAEXCECAO = 'E'  										 										"
		vCriterio = vCriterio + "									AND X.PRESTADOR    = PE.PRESTADOR  								 										"
		vCriterio = vCriterio + "									AND X.DATAINICIAL <= "+SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)+"                  	"
		vCriterio = vCriterio + "									AND (X.DATAFINAL IS NULL                                         										"
		vCriterio = vCriterio + "									 OR  X.DATAFINAL  >= "+SQLDate(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)+"))                  	"
		vCriterio = vCriterio + "								  UNION  															 										"
		vCriterio = vCriterio + "								 SELECT X.EVENTO 													 										"
		vCriterio = vCriterio + "								   FROM SAM_PRESTADOR_REGRA X 										 										"
		vCriterio = vCriterio + "								  WHERE X.REGRAEXCECAO = 'R' 										 										"
		vCriterio = vCriterio + "									AND X.PRESTADOR	   = "+CurrentQuery.FieldByName("PRESTADOR").AsString+"							 		"
		vCriterio = vCriterio + "									AND X.DATAINICIAL <= "+SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)+"                  	"
		vCriterio = vCriterio + "									AND (X.DATAFINAL IS NULL                                         										"
		vCriterio = vCriterio + "									 OR  X.DATAFINAL  >= "+SQLDate(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)+")                 	"
		vCriterio = vCriterio + ")																																			"
	Else
		vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
	End If
  criarCriterio = vCriterio

End Function
