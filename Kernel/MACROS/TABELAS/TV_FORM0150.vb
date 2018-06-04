'HASH: 816EC5D8E868B0579B7629392B8752F2
'#Uses "*bsShowMessage"
'#Uses "*VerificarPegReapresentadoExercicioPosterior"
'#Uses "*RecordHandleOfTableInterfacePEG"
'#Uses "*RefreshNodesWithTableInterfacePEG"

Public Sub EMPENHO_OnPopup(ShowPopup As Boolean)
    Dim interface As Object
	Dim vHandleEmpenho As Long
	Dim vColunas As String
	Dim vCampos As String
	Dim vCriterio As String
	Dim vtabelas As String
    Dim qRecebedor As BPesquisa

    Set qRecebedor = NewQuery

    qRecebedor.Clear
    qRecebedor.Add("SELECT RECEBEDOR FROM SAM_PEG WHERE HANDLE = :HANDLE")
    qRecebedor.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    qRecebedor.Active = True

    Set interface = CreateBennerObject("Procura.Procurar")

    vColunas = "SFN_EMPENHO.NUMERO|SFN_EMPENHO.DESCRICAO|SFN_DOTACAOEXERCICIO.EXERCICIO|UGRESPONSAVEL.DESCRICAO|NATUREZADESPESA.DESCRICAO"

    vCampos = "Número|Empenho|Exercício|Unidade Gestora|Natureza da Despesa"

	vCriterio = "(SFN_EMPENHO.TABTIPO = 1) OR (SFN_EMPENHO.TABTIPO = 2 AND SFN_EMPENHO.PRESTADOR = " + CStr(qRecebedor.FieldByName("RECEBEDOR").AsInteger) + ")"

	vtabelas = "SFN_EMPENHO|SFN_DOTACAONATUREZA[SFN_DOTACAONATUREZA.HANDLE = SFN_EMPENHO.DOTACAONATUREZA]|SFN_DOTACAO[SFN_DOTACAO.HANDLE = SFN_DOTACAONATUREZA.DOTACAO]|SFN_DOTACAOEXERCICIO[SFN_DOTACAOEXERCICIO.HANDLE = SFN_DOTACAO.EXERCICIO]|NATUREZADESPESA[NATUREZADESPESA.HANDLE = SFN_DOTACAONATUREZA.NATUREZADESPESA]|UGRESPONSAVEL[UGRESPONSAVEL.HANDLE = SFN_DOTACAO.UGRESPONSAVEL]"

	vHandleEmpenho = interface.Exec(CurrentSystem, vtabelas, vColunas, 2, vCampos, vCriterio, "Empenho", True, "", EMPENHO.LocateText)

    Set qRecebedor = Nothing

    If (vHandleEmpenho <> 0) Then
        CurrentQuery.FieldByName("EMPENHO").AsInteger = vHandleEmpenho
    End If

    ShowPopup = False
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If Not VerificarPegReapresentadoExercicioPosterior(RecordHandleOfTableInterfacePEG("SAM_PEG")) Then
      bsShowMessage("O Empenho não pode ser alterado. ", "E")
      CanContinue = False
	  Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

    Dim viHandlePeg As Long
	Dim qSql As BPesquisa
	Dim qUpdate As BPesquisa

	Set qSql = NewQuery
	Set qUpdate = NewQuery

	If Not CurrentQuery.FieldByName("EMPENHO").IsNull Then
	    qSql.Clear
		qSql.Add("SELECT E.TABTIPO    TABTIPO,                                   ")
		qSql.Add("       E.PRESTADOR  PRESTADOR,                                 ")
		qSql.Add("       N.HANDLE     HANDLENATUREZA,                            ")
		qSql.Add("       D.HANDLE     HANDLEDOTACAO,                             ")
        qSql.Add("       EX.HANDLE    HANDLEEXERCICIO                            ")
        qSql.Add("  FROM SFN_EMPENHO E                                           ")
        qSql.Add("  JOIN SFN_DOTACAONATUREZA N ON E.DOTACAONATUREZA = N.HANDLE   ")
        qSql.Add("  JOIN SFN_DOTACAO D ON N.DOTACAO = D.HANDLE                   ")
        qSql.Add("  JOIN SFN_DOTACAOEXERCICIO EX ON D.EXERCICIO = EX.HANDLE      ")
		qSql.Add(" WHERE E.HANDLE = :HANDLEEMPENHO                               ")
		qSql.ParamByName("HANDLEEMPENHO").AsInteger = CurrentQuery.FieldByName("EMPENHO").AsInteger
		qSql.Active = True

		If WebMode Then

			Dim qPrestadordor As BPesquisa
   			Set qPrestadordor = NewQuery

			qPrestadordor.Clear
		    qPrestadordor.Add("SELECT RECEBEDOR FROM SAM_PEG WHERE HANDLE = :HANDLE")
		    qPrestadordor.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
		    qPrestadordor.Active = True

			If (qSql.FieldByName("TABTIPO").AsInteger = 2) And (qSql.FieldByName("PRESTADOR").AsInteger <> qPrestadordor.FieldByName("RECEBEDOR").AsInteger) Then
				bsShowMessage("Informado Empenho Orçamentário incompatível com Recebedor","E")
				CanContinue = False
				Exit Sub
			End If
			Set qPrestadordor = Nothing
		End If

        qUpdate.Clear
        qUpdate.Add(" UPDATE SAM_PEG                                 ")
        qUpdate.Add("    SET EMPENHOPEG = :HANDLEEMPENHO,            ")
        qUpdate.Add("        DOTACAONATUREZAPEG = :HANDLENATUREZA,   ")
		qUpdate.Add("        DOTACAOPEG = :HANDLEDOTACAO,            ")
		qUpdate.Add("        DOTACAOEXERCICIOPEG = :HANDLEEXERCICIO  ")
		qUpdate.Add("  WHERE HANDLE = :HANDLEPEG                     ")
		qUpdate.ParamByName("HANDLEEMPENHO").AsInteger = CurrentQuery.FieldByName("EMPENHO").AsInteger
		qUpdate.ParamByName("HANDLENATUREZA").AsInteger = qSql.FieldByName("HANDLENATUREZA").AsInteger
		qUpdate.ParamByName("HANDLEDOTACAO").AsInteger = qSql.FieldByName("HANDLEDOTACAO").AsInteger
		qUpdate.ParamByName("HANDLEEXERCICIO").AsInteger = qSql.FieldByName("HANDLEEXERCICIO").AsInteger
		qUpdate.ParamByName("HANDLEPEG").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
		qUpdate.ExecSQL

        bsShowMessage("Alteração Concluída","I")

		RefreshNodesWithTableInterfacePEG("SAM_PEG")
	End If

	Set qUpdate = Nothing
	Set qSql = Nothing
End Sub
