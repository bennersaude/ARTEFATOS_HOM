'HASH: DB990352678A7B2E3A2C3518A5D61E83
'Macro: SAM_TGE_TABELATISS

'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
  Dim qSincroniza As Object
  Set qSincroniza = NewQuery

  qSincroniza.Add("SELECT HANDLE, DESCRICAO, CLASSEEVENTO")
  qSincroniza.Add("  FROM SAM_TGE                        ")
  qSincroniza.Add(" WHERE HANDLE = :HANDLETGE            ")
  qSincroniza.ParamByName("HANDLETGE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  qSincroniza.Active = True

  Dim callEntity As CSEntityCall

  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamTge, Benner.Saude.Entidades", "EnviaProcedimentoParaSincronizacaoIntegracaoHospitalar")
  callEntity.AddParameter(pdtInteger, qSincroniza.FieldByName("HANDLE").AsInteger)
  callEntity.AddParameter(pdtString, "Descricao Antiga")
  callEntity.AddParameter(pdtString, qSincroniza.FieldByName("DESCRICAO").AsString)
  callEntity.AddParameter(pdtAutomatic, 0)
  callEntity.AddParameter(pdtAutomatic, IIf(IsNull(qSincroniza.FieldByName("CLASSEEVENTO").AsInteger),0,qSincroniza.FieldByName("CLASSEEVENTO").AsInteger))

  callEntity.Execute

  Set callEntity = Nothing
  Set qSincroniza = Nothing
End Sub

Public Sub TABLE_AfterScroll()

	If VisibleMode Then
		TABELATISS.LocalWhere = " TIS_TABELAPRECO.HANDLE NOT IN (SELECT TABELATISS        " + _
								"                  FROM SAM_TGE_TABELATISS                " + _
								"                 WHERE EVENTO = @EVENTO)                 "
	Else
		TABELATISS.WebLocalWhere = " A.HANDLE NOT IN (SELECT TABELATISS                   " + _
							       "                  FROM SAM_TGE_TABELATISS             " + _
							       "                 WHERE EVENTO = @CAMPO(EVENTO))       "
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSQL As Object
  Set qSQL = NewQuery

  qSQL.Add("SELECT TGE.HANDLE, TGE.ESTRUTURA, TGE.DESCRICAO, M.DESCRICAO MASCARA")
  qSQL.Add("  FROM SAM_TGE TGE")
  qSQL.Add("  JOIN SAM_TGE_TABELATISS T ON (T.EVENTO = TGE.HANDLE)")
  qSQL.Add("  JOIN SAM_MASCARATGE M ON (M.HANDLE = TGE.MASCARATGE)")
  qSQL.Add(" WHERE TGE.HANDLE <> :HANDLETGE")
  qSQL.Add("   AND TGE.ESTRUTURANUMERICA = (SELECT TGE2.ESTRUTURANUMERICA")
  qSQL.Add("                                  FROM SAM_TGE TGE2")
  qSQL.Add("                                 WHERE TGE2.HANDLE = :HANDLETGE)")
  qSQL.Add("   AND T.TABELATISS = :HANDLETABELATISS")
  qSQL.ParamByName("HANDLETGE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  qSQL.ParamByName("HANDLETABELATISS").AsInteger = CurrentQuery.FieldByName("TABELATISS").AsInteger
  qSQL.Active = True

  If qSQL.FieldByName("HANDLE").AsInteger > 0 Then
    Dim vsMensagem As String

    vsMensagem = "Tabela informada não pode ser cadastrada pois já existe a mesma combinação de estrutura de evento e tabela TISS cadastrada." + Chr(13) + Chr(13) + _
                 "Ocorrência(s) encontrada(s):"

    While Not qSQL.EOF
      vsMensagem = vsMensagem + Chr(13) + qSQL.FieldByName("MASCARA").AsString + ": " + qSQL.FieldByName("ESTRUTURA").AsString + " - " + qSQL.FieldByName("DESCRICAO").AsString
      qSQL.Next
    Wend

    Dim qTABELAPRECO As Object
  	Set qTABELAPRECO = NewQuery

	qTABELAPRECO.Add("Select COUNT(*) QTDE ")
  	qTABELAPRECO.Add("  FROM TIS_TABELAPRECO")
 	qTABELAPRECO.Add(" WHERE HANDLE = :HANDLETABELATISS")
   	qTABELAPRECO.Add("   And CODIGO = '00' ")
   	qTABELAPRECO.ParamByName("HANDLETABELATISS").AsInteger = CurrentQuery.FieldByName("TABELATISS").AsInteger
   	qTABELAPRECO.Active = True

   	If qTABELAPRECO.FieldByName("QTDE").AsInteger > 0 Then
   		If (WebMode) Then
			If bsShowMessage(vsMensagem + Chr(13) + Chr(13) +"Deseja continuar ?", "Q") = vbYes Then
				CanContinue = True
	   		End If
	   	Else
	   		If bsShowMessage(vsMensagem + Chr(13) + Chr(13) +"Deseja continuar ?", "Q") = vbNo Then
				CanContinue = False
	   		End If
   		End If
   	Else
		bsShowMessage(vsMensagem, "E")
		CanContinue = False
   	End If

	Set qTABELAPRECO = Nothing

  End If

  Set qSQL = Nothing
End Sub
