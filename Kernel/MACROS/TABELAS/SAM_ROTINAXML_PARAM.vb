'HASH: 12AB8459FC4198022AC8D2F4A747FEAD
'Macro da tabela: SAM_ROTINAXML_PARAM
'#Uses "*bsShowMessage"

Option Explicit

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)

    If (CurrentQuery.FieldByName("ROTINAFIN").IsNull) Then
        bsShowMessage("Informe a rotina financeira.", "I")
        ShowPopup = False
        Exit Sub
    End If

    Dim qRotinafin As BPesquisa
    Set qRotinafin = NewQuery

    qRotinafin.Clear
    qRotinafin.Add("SELECT FFP.TABFATURAR TABFATURAR,								   ")
    qRotinafin.Add("		 FFP.CONVENIO CONVENIO,									   ")
    qRotinafin.Add("		 FFP.GRUPOCONTRATO GRUPOCONTRATO,						   ")
    qRotinafin.Add("		 FFP.CONTRATOINICIAL INICIAL,							   ")
    qRotinafin.Add("		 FFP.CONTRATOFINAL FINAL,								   ")
    qRotinafin.Add("		 FFP.CONTRATO CONTRATO									   ")
    qRotinafin.Add("  FROM SFN_ROTINAFINFAT_PARAM FFP								   ")
    qRotinafin.Add("  JOIN SFN_ROTINAFINFAT       FF  ON (FF.HANDLE = FFP.ROTINAFINFAT)")
    qRotinafin.Add("  JOIN SFN_ROTINAFIN          RF  ON (RF.HANDLE = FF.ROTINAFIN)    ")
    qRotinafin.Add(" WHERE RF.HANDLE = :HANDLE                                         ")
    qRotinafin.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    qRotinafin.Active = True

	Dim vsCriterio As String

	vsCriterio = "A.CONVENIORECIPROCIDADE = 'S' "

	If (qRotinafin.EOF) Then

	    bsShowMessage("Esta rotina financeira não possui filtro para contratos, portanto, não poderá ter guias exportadas.", "I")
        ShowPopup = False
        Exit Sub

	Else

		vsCriterio = vsCriterio + "AND ("

		While Not qRotinafin.EOF
			Select Case qRotinafin.FieldByName("TABFATURAR").AsInteger
				Case 1
					vsCriterio = vsCriterio + " A.CONVENIO = " + qRotinafin.FieldByName("CONVENIO").AsString + " OR"
				Case 2
					vsCriterio = vsCriterio + " A.GRUPOCONTRATO = " + qRotinafin.FieldByName("GRUPOCONTRATO").AsString + " OR"
				Case 3
					vsCriterio = vsCriterio + " A.HANDLE BETWEEN " + qRotinafin.FieldByName("INICIAL").AsString + " AND " + qRotinafin.FieldByName("FINAL").AsString + " OR"
			End Select

			qRotinafin.Next
		Wend

		vsCriterio = Left(vsCriterio, Len(vsCriterio) - 3) + ") "
	End If

	CONTRATO.LocalWhere = vsCriterio

	qRotinafin.Active = False
    Set qRotinafin = Nothing

End Sub

Public Sub ROTINAFIN_OnChange()

    CurrentQuery.FieldByName("CONTRATO").Clear

End Sub

Public Sub ROTINAFIN_OnPopup(ShowPopup As Boolean)

    Dim qCompet As BPesquisa
	Set qCompet = NewQuery

    qCompet.Clear
	qCompet.Add("SELECT HANDLE													 ")
	qCompet.Add("  FROM SFN_COMPETFIN											 ")
	qCompet.Add(" WHERE COMPETENCIA = (SELECT COMPETENCIA						 ")
	qCompet.Add("                        FROM SAM_COMPETXML						 ")
	qCompet.Add("                       WHERE HANDLE = (SELECT COMPETENCIA		 ")
	qCompet.Add("                                         FROM SAM_ROTINAXML	 ")
	qCompet.Add("                                        WHERE HANDLE = :HANDLE))")
	qCompet.Add("   AND TIPOFATURAMENTO = (SELECT HANDLE						 ")
	qCompet.Add("							  FROM SIS_TIPOFATURAMENTO			 ")
	qCompet.Add("							 WHERE CODIGO = :CODIGO)			 ")
	qCompet.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAXML").AsInteger
	qCompet.ParamByName("CODIGO").AsInteger = 110
	qCompet.Active = True


	If (qCompet.FieldByName("HANDLE").AsInteger > 0) Then
		ROTINAFIN.LocalWhere = "A.TIPOFATURAMENTO = (SELECT HANDLE FROM SIS_TIPOFATURAMENTO WHERE CODIGO = 110) AND A.COMPETFIN = " + qCompet.FieldByName("HANDLE").AsString
	Else
		ShowPopup = False
		bsShowMessage("Não existe rotina financeira para esta competência.", "I")
	End If

    qCompet.Active = False
	Set qCompet = Nothing

End Sub

Public Sub TABLE_AfterScroll()

    Dim qOperadoras As BPesquisa
    Set qOperadoras = NewQuery

    qOperadoras.Clear
    qOperadoras.Add("SELECT COUNT(1) QTD                                            ")
    qOperadoras.Add("  FROM SAM_OPERADORA                                           ")
    qOperadoras.Add(" WHERE :HOJE BETWEEN DATAINICIAL AND COALESCE(DATAFINAL, :HOJE)")
    qOperadoras.ParamByName("HOJE").AsDateTime = ServerDate
    qOperadoras.Active = True

    OPERADORA.ReadOnly = qOperadoras.FieldByName("QTD").AsInteger = 1

    qOperadoras.Active = False
    Set qOperadoras = Nothing

    BOTAOCANCELAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString <> "1")
    BOTAOPROCESSAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "1")
    BOTAOEXPORTAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "5") 'Processada.

End Sub
