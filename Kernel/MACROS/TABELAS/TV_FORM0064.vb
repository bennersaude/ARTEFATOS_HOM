'HASH: A684AF5ECF0DD62F8EB0C34D8FD1B08B
 '#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	RELATORIO.WebLocalWhere = " NOME LIKE 'CAR%' AND (NOME LIKE '% DEFERIDO%' OR NOME LIKE '% INDEFERIDO%' OR NOME LIKE '% ANALISE%' OR NOME LIKE '% DEVOLVIDO%' OR NOME LIKE '% PARCIAL%')"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	'SMS 98377 Bruno Penteado
	  Dim sqlRel As Object
	  Set sqlRel = NewQuery
	  sqlRel.Add("SELECT CODIGO, NOME FROM R_RELATORIOS WHERE HANDLE = :HANDLE")
	  sqlRel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("RELATORIO").AsInteger
	  sqlRel.Active = True

	  Dim qUpd As Object
	  Set qUpd = NewQuery

	  qUpd.Add("UPDATE SAM_TIPOPROCESSOCREDENCTO SET")


	  Select Case CurrentQuery.FieldByName("TIPORELATORIO").AsString
	    Case "D"
		  If InStr(sqlRel.FieldByName("NOME").AsString, "DEFERIDO") > 0 Then
	      	qUpd.Add(" RELATORIODEFERIDO = :CODIGO")
	      Else
			bsShowMessage("O relatório não é do tipo deferido", "E")
			CanContinue = False
			Exit Sub
	      End If

	    Case "I"
		  If InStr(sqlRel.FieldByName("NOME").AsString, "INDEFERIDO") > 0 Then
	      	qUpd.Add(" RELATORIOINDEFERIDO = :CODIGO")
	      Else
			bsShowMessage("O relatório não é do tipo indeferido", "E")
			CanContinue = False
			Exit Sub
	      End If

	    Case "A"
	      If InStr(sqlRel.FieldByName("NOME").AsString, "ANALISE") > 0 Then
	      	qUpd.Add(" RELATORIOEMANALISE = :CODIGO")
	      Else
			bsShowMessage("O relatório não é do tipo análise", "E")
			CanContinue = False
			Exit Sub
	      End If

	    Case "V"
	      If InStr(sqlRel.FieldByName("NOME").AsString, "DEVOLVIDO") > 0 Then
	      	qUpd.Add(" RELATORIODEVOLVIDO = :CODIGO")
	      Else
			bsShowMessage("O relatório não é do tipo devolvido", "E")
			CanContinue = False
			Exit Sub
	      End If

	    Case "P"
	      If InStr(sqlRel.FieldByName("NOME").AsString, "PARCIAL") > 0 Then
	      	qUpd.Add(" RELATORIOPARCIAL = :CODIGO")
	      Else
			bsShowMessage("O relatório não é do tipo parcial", "E")
			CanContinue = False
			Exit Sub
	      End If

	  End Select


	  qUpd.Add(" WHERE HANDLE  = :HANDLE")


	  qUpd.ParamByName("CODIGO").AsString = sqlRel.FieldByName("CODIGO").AsString
	  qUpd.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_TIPOPROCESSOCREDENCTO")


	  qUpd.ExecSQL

	  Set qUpd = Nothing
	  Set sqlRel = Nothing
End Sub
