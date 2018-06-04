'HASH: 6E5F8000BAAE8A894E1305C528EF18E7
'Macro: SAM_PRECOPRESTADOR_DOTAC
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*NegociacaoPrecos"

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
	Dim SQL,SQL2 As Object
	Dim Nivel As Integer
	Dim result As String

	Set SQL =NewQuery

	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active =True

	Nivel =-1

	If CurrentQuery.FieldByName("EVENTO").IsNull Then
		result =""
	Else
		Set SQL2 =NewQuery

		SQL2.Add("SELECT ASSOCIACAO FROM SAM_PRESTADOR WHERE HANDLE = " +CurrentQuery.FieldByName("PRESTADOR").AsString +"")

		SQL2.Active =True

		If SQL2.FieldByName("ASSOCIACAO").AsString ="S" Then
			If SQL.FieldByName("NIVEL1").AsInteger =5 Then
				Nivel =1
			ElseIf SQL.FieldByName("NIVEL2").AsInteger =5 Then
				Nivel =2
			ElseIf SQL.FieldByName("NIVEL3").AsInteger =5 Then
				Nivel =3
			ElseIf SQL.FieldByName("NIVEL4").AsInteger =5 Then
				Nivel =4
			ElseIf SQL.FieldByName("NIVEL5").AsInteger =5 Then
				Nivel =5
			ElseIf SQL.FieldByName("NIVEL6").AsInteger =5 Then
				Nivel =6
			ElseIf SQL.FieldByName("NIVEL7").AsInteger =5 Then
				Nivel =7
			ElseIf SQL.FieldByName("NIVEL8").AsInteger =5 Then
				Nivel =8
			End If

			If Nivel <>-1 Then
				Set Interface =CreateBennerObject("BSPRE001.Rotinas")

				ValorEvento =Interface.ValorEvento(CurrentSystem,vDataBaseChecagemVigencia,99,-1,-1,CurrentQuery.FieldByName("PRESTADOR").Value,-1,-1,-1,-1,-1,CurrentQuery.FieldByName("EVENTO").Value,-1,Nivel,CurrentQuery.FieldByName("CONVENIO").AsInteger,CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString, CurrentQuery.FieldByName("CBOS").AsString)

				result ="Valor do evento nesta vigência: R$ " +Format(ValorEvento,"#,##0.0000")+" ("+Format(ValorEvento,"#,##0.00")+")"
			Else
				result ="Na configuração de busca de preço, não foi definido um nível para a Associação!"
			End If
		Else
			If SQL.FieldByName("NIVEL1").AsInteger =4 Then
				Nivel =1
			ElseIf SQL.FieldByName("NIVEL2").AsInteger =4 Then
				Nivel =2
			ElseIf SQL.FieldByName("NIVEL3").AsInteger =4 Then
				Nivel =3
			ElseIf SQL.FieldByName("NIVEL4").AsInteger =4 Then
				Nivel =4
			ElseIf SQL.FieldByName("NIVEL5").AsInteger =4 Then
				Nivel =5
			ElseIf SQL.FieldByName("NIVEL6").AsInteger =4 Then
				Nivel =6
			ElseIf SQL.FieldByName("NIVEL7").AsInteger =4 Then
				Nivel =7
			ElseIf SQL.FieldByName("NIVEL8").AsInteger =4 Then
				Nivel =8
			End If

			If Nivel <>-1 Then
				Set Interface =CreateBennerObject("BSPRE001.Rotinas")

				ValorEvento =Interface.ValorEvento(CurrentSystem,vDataBaseChecagemVigencia,99,-1,-1,CurrentQuery.FieldByName("PRESTADOR").Value,-1,-1,-1,-1,-1,CurrentQuery.FieldByName("EVENTO").Value,-1,Nivel,CurrentQuery.FieldByName("CONVENIO").AsInteger,CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString, CurrentQuery.FieldByName("CBOS").AsString)

				result ="Valor do evento nesta vigência: R$ " +Format(ValorEvento,"###,###,##0.0000")+" ("+Format(ValorEvento,"#,##0.00")+")"
			Else
				result ="Na configuração de busca de preço, não foi definido um nível para o Prestador!"
			End If
		End If
	End If

	If VisibleMode Then
		LABELPRECO.Text = result
	Else
		If Nivel <> -1 Then
			bsShowMessage(result,"I")
		Else
			bsShowMessage(result, "E")
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

	If CurrentQuery.State <>3 Then
		Exit Sub
	End If

	If CurrentQuery.FieldByName("EVENTO").IsNull Then
		bsShowMessage("Digite um evento !", "I")
		Exit Sub
	End If

	Set Interface =CreateBennerObject("Procura.Procurar")

	vColunas ="SAM_PRECOGENERICO.DESCRICAO|SAM_TGE.DESCRICAO|SAM_PRECOGENERICO_DOTAC.QTDUSHONORARIO|SAM_PRECOGENERICO_DOTAC.QTDUSCUSTOOPERACIONAL|SAM_PRECOGENERICO_DOTAC.FATORFILME|SAM_PRECOGENERICO_DOTAC.PORTEANESTESICO|SAM_PRECOGENERICO_DOTAC.PORTESALA"
	vCriterio ="((SAM_PRECOGENERICO_DOTAC.DATAINICIAL <= " +SQLDate(ServerDate) +") AND (SAM_PRECOGENERICO_DOTAC.DATAFINAL >= " +SQLDate(ServerDate) +" OR SAM_PRECOGENERICO_DOTAC.DATAFINAL IS NULL))"
	vCriterio =vCriterio +" AND SAM_PRECOGENERICO_DOTAC.EVENTO = " +CurrentQuery.FieldByName("EVENTO").AsString
	vCampos ="Tabela Genérica|Evento|Qtde US Honorário|Qtde US Custo Operacional|Fator de filme|Porte Anestésico|Porte de Sala"
	vHandle =Interface.Exec(CurrentSystem,"SAM_PRECOGENERICO_DOTAC|SAM_PRECOGENERICO[SAM_PRECOGENERICO_DOTAC.TABELAPRECO = SAM_PRECOGENERICO.HANDLE]|SAM_TGE[SAM_PRECOGENERICO_DOTAC.EVENTO=SAM_TGE.HANDLE]",vColunas,1,vCampos,vCriterio,"Quantidades",True,"","")

	Set QueryRetorno =NewQuery

	QueryRetorno.Add("SELECT * FROM SAM_PRECOGENERICO_DOTAC WHERE HANDLE =:HANDLE")
	QueryRetorno.ParamByName("HANDLE").Value =vHandle
	QueryRetorno.Active =True

	CurrentQuery.FieldByName("QTDUSHONORARIO").Value =QueryRetorno.FieldByName("QTDUSHONORARIO").Value
	CurrentQuery.FieldByName("QTDUSCUSTOOPERACIONAL").Value =QueryRetorno.FieldByName("QTDUSCUSTOOPERACIONAL").Value
	CurrentQuery.FieldByName("FATORFILME").Value =QueryRetorno.FieldByName("FATORFILME").Value
	CurrentQuery.FieldByName("PORTEANESTESICO").Value =QueryRetorno.FieldByName("PORTEANESTESICO").Value
	CurrentQuery.FieldByName("PORTESALA").Value =QueryRetorno.FieldByName("PORTESALA").Value

	QueryRetorno.Active =False

	Set Interface =Nothing
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

  SessionVar("TabelaDeDotacao") = "SAM_PRECOPRESTADOR_DOTAC"
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

  RefreshNodesWithTable("SAM_PRECOPRESTADOR_DOTAC")
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim vData As String
	Dim Interface As Object
	Dim vColunas,vCampos,vTabela As String

	Set Interface =CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vHandle = Interface.Exec(CurrentSystem,vTabela,vColunas,1,vCampos,CriarCriterio,"Eventos que o prestador pode executar",True,EVENTO.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value =vHandle
	End If

	Set Interface = Nothing
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

	ShowPopup =False
	vHandle =ProcuraTabelaFilme(TABELAFILME.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAFILME").Value =vHandle
	End If
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup =False
	vHandle =ProcuraTabelaUS(TABELAUS.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAUS").Value =vHandle
	End If
End Sub

Public Sub TABLE_AfterEdit()
	UpdateLastUpdate("SAM_CONVENIO")

	Dim vCondicao As String

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere =vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
		EVENTO.WebLocalWhere   = CriarCriterio
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qCBOS            As Object
	Dim vFiltroAdicional As String
	Dim vAtedias As Integer
	Dim vDeDias As Integer
	Dim vAteAnos As Integer
	Dim vDeAnos As Integer

	'SMS 49152 - Anderson Lonardoni
	'Esta verificação foi tirada do BeforeInsert e colocada no
	'BeforePost para que, no caso de Inserção, já existam valores
	'no CurrentQuery e para funcionar com o Integrator

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

	Dim Msg As String

    If verificaChecaFilial Then
  	  If checkPermissaoFilial(CurrentSystem,"I","P",Msg)="N" Then
		bsShowMessage(Msg, "E")
		CanContinue =False
		Exit Sub
	  End If
	End If
	'SMS 49152 - Fim

    If CurrentQuery.FieldByName("CBOS").IsNull Then
    	vFiltroAdicional = " AND (CBOS IS NULL OR CBOS = '')"
	Else
		vFiltroAdicional = " AND CBOS = '" + CurrentQuery.FieldByName("CBOS").AsString + "'"
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

	CanContinue = ValidacoesBeforePostNegociacaoPreco(CurrentQuery.FieldByName("HANDLE").AsInteger, "SAM_PRECOPRESTADOR_DOTAC", "DATAINICIAL", "DATAFINAL", "PRESTADOR", _
	  CurrentQuery.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("EVENTO").AsInteger, CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString, _
	  CurrentQuery.FieldByName("CONVENIO").AsString, vFiltroAdicional, vDeAnos, vDeDias, _
	  vAteAnos, vAtedias, CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger, _
	  CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

    If Not CanContinue Then
      Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

    If verificaChecaFilial Then
      If checkPermissaoFilial(CurrentSystem,"E","P",Msg)="N" Then
		bsShowMessage(Msg, "E")
		CanContinue =False
		Exit Sub
	  End If
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

    If verificaChecaFilial Then
  	  If checkPermissaoFilial(CurrentSystem,"A","P",Msg)="N" Then
		bsShowMessage(Msg, "E")
		CanContinue =False
		Exit Sub
	  End If
	End If
End Sub

Public Sub TABLE_AfterScroll()
	LABELPRECO.Text =""
End Sub

Public Sub TABLE_AfterInsert()
	Dim SQL As Object
	Set SQL =NewQuery

	SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

	SQL.Active =True

	If SQL.FieldByName("TOTAL").AsInteger =1 Then
		SQL.Active =False

		SQL.Clear

		SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

		SQL.Active =True

		CurrentQuery.FieldByName("CONVENIO").Value =SQL.FieldByName("HANDLE").Value
	End If

	Set SQL =Nothing

	UpdateLastUpdate("SAM_CONVENIO")

	Dim vCondicao As String

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere =vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
		EVENTO.WebLocalWhere   = CriarCriterio
	End If
End Sub

Public Function CriarCriterio As String
	Dim qPrestador As BPesquisa
	Dim vCriterio  As String
	Dim vDataFinalAux As Date
	Dim vDataInicialAux As Date
	Dim vDataInicial As String
	Dim vDataFinal  As String

	Set qPrestador = NewQuery

	qPrestador.Active = False

	qPrestador.Add("SELECT ASSOCIACAO FROM SAM_PRESTADOR WHERE HANDLE = :PRESTADOR")

	qPrestador.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

	qPrestador.Active = True

	If (qPrestador.FieldByName("ASSOCIACAO").AsString <> "S") Then
		If VisibleMode Then
			vCriterio = "(SAM_TGE.HANDLE IN ( SELECT DISTINCT GE.EVENTO"
		Else
			vCriterio = "((A.HANDLE IN ( SELECT DISTINCT GE.EVENTO"
		End If

 	   If CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
 	     vDataInicialAux = ServerDate
 	   Else
	     vDataInicialAux = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
	   End If

	   If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
	     vDataFinalAux = ServerDate
	   Else
	     vDataFinalAux = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
	   End If

	   If (vDataFinalAux < vDataInicialAux) Then
	     vDataFinalAux = vDataInicialAux
	   End If

	   vDataInicial = SQLDate(vDataInicialAux)
	   vDataFinal = SQLDate(vDataFinalAux)

		vCriterio = vCriterio +    " FROM SAM_ESPECIALIDADEGRUPO_EXEC    GE  "
		vCriterio = vCriterio +    " JOIN SAM_ESPECIALIDADEGRUPO         EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO)  "
		vCriterio = vCriterio +    " JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE = EG.ESPECIALIDADE)  "
		vCriterio = vCriterio +    " JOIN SAM_PRESTADOR_ESPECIALIDADE    PE ON (PE.ESPECIALIDADE = E.HANDLE)  "
		vCriterio = vCriterio +    " LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)  "
		vCriterio = vCriterio + " WHERE PE.DATAINICIAL <= " + vDataInicial
		vCriterio = vCriterio +   " AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= " + vDataFinal + ")"

		If VisibleMode Then
			vCriterio = vCriterio +   " AND PE.PRESTADOR = " + CStr(CurrentQuery.FieldByName("PRESTADOR").AsInteger)
		Else
			vCriterio = vCriterio +   " AND PE.PRESTADOR = @CAMPO(PRESTADOR)"
		End If

		vCriterio = vCriterio +   " AND GE.EVENTO NOT IN (SELECT X.EVENTO  "
		vCriterio = vCriterio +                           " FROM SAM_PRESTADOR_REGRA X  "
		vCriterio = vCriterio +                          " WHERE X.REGRAEXCECAO = 'E' AND X.PERMITERECEBER = 'S' "
		vCriterio = vCriterio +                            " AND X.PRESTADOR = PE.PRESTADOR  "
		vCriterio = vCriterio +                            " AND X.DATAINICIAL <= " + vDataInicial
		vCriterio = vCriterio +                            " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vDataFinal + ")))"
		If VisibleMode Then
		  vCriterio = vCriterio + " OR SAM_TGE.HANDLE IN(  "
		Else
          vCriterio = vCriterio + " OR A.HANDLE IN(  "
		End If
		vCriterio = vCriterio +  " SELECT X.EVENTO "
		vCriterio = vCriterio +    " FROM SAM_PRESTADOR_REGRA X "
		vCriterio = vCriterio +    " WHERE X.REGRAEXCECAO = 'R' AND X.PERMITERECEBER = 'S' "

		If VisibleMode Then
			vCriterio = vCriterio + " AND X.PRESTADOR = " + CStr(CurrentQuery.FieldByName("PRESTADOR").AsInteger)
		Else
			vCriterio = vCriterio + " AND X.PRESTADOR = @CAMPO(PRESTADOR)"
		End If

		vCriterio = vCriterio +      " AND X.DATAINICIAL <= " + vDataInicial
		vCriterio = vCriterio +      " AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vDataFinal + ")"
		vCriterio = vCriterio + " )) "

        If WebMode Then
          vCriterio = vCriterio + "OR (A.ULTIMONIVEL = 'N'))"
        End If
	Else
		If VisibleMode Then
			vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
		Else
			vCriterio = "A.ULTIMONIVEL = 'S'"
		End If
	End If
	Set qPrestador = Nothing
	CriarCriterio = vCriterio
End Function

Public Function verificaChecaFilial As Boolean

  If (Not VisibleMode) And WebMode Then
    Dim qParamPrestador As Object
    Set qParamPrestador = NewQuery

    qParamPrestador.Clear
    qParamPrestador.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
    qParamPrestador.Active = True
    If qParamPrestador.FieldByName("CONTROLEDEACESSO").AsString = "N" Then
      verificaChecaFilial = False
    Else
      verificaChecaFilial = True
    End If

    Set qParamPrestador = Nothing

  ElseIf (Not VisibleMode) And (Not WebMode) Then
    verificaChecaFilial = False
  Else
    verificaChecaFilial = True
  End If

End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOPRECO"
			BOTAOPRECO_OnClick
		Case "BOTAOQUANTIDADES"
			BOTAOQUANTIDADES_OnClick
		Case "BOTAOREPLICARPORCBOS"
            SessionVar("TabelaDeDotacao") = "SAM_PRECOPRESTADOR_DOTAC"
            SessionVar("HandleDotac") = CurrentQuery.FieldByName("HANDLE").AsString
	End Select
End Sub
