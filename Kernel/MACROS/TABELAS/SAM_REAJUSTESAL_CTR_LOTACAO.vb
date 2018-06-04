'HASH: FC6AC7E04E54BC097175B830942DB304
'#Uses "*bsShowMessage"

Function RotinaGerada As Boolean
	Dim Query As Object
	Set Query = NewQuery

	Query.Clear
	Query.Add("SELECT SITUACAOGERACAO")
	Query.Add("  FROM SAM_REAJUSTESAL_PARAM")
	Query.Add(" WHERE HANDLE = :HANDLE")
	Query.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_REAJUSTESAL_PARAM")
	Query.Active = True

	RotinaGerada = Query.FieldByName("SITUACAOGERACAO").AsInteger > 1

	Set Query = Nothing
End Function

Public Sub TABLE_AfterScroll()
	If WebMode Then
		LOTACAO.WebLocalWhere = " CONTRATO = (SELECT CONTRATO FROM SAM_REAJUSTESAL_CTR WHERE HANDLE = " & RecordHandleOfTable("SAM_REAJUSTESAL_CTR") & ")" & _
		                        " AND HANDLE NOT IN (SELECT LOTACAO FROM SAM_REAJUSTESAL_CTR_LOTACAO WHERE REAJUSTESALCTR = " &  RecordHandleOfTable("SAM_REAJUSTESAL_CTR") & ")"
	Else
		LOTACAO.LocalWhere = " CONTRATO = (SELECT CONTRATO FROM SAM_REAJUSTESAL_CTR WHERE HANDLE = " & RecordHandleOfTable("SAM_REAJUSTESAL_CTR") & ")" & _
		                        " AND HANDLE NOT IN (SELECT LOTACAO FROM SAM_REAJUSTESAL_CTR_LOTACAO WHERE REAJUSTESALCTR = " &  RecordHandleOfTable("SAM_REAJUSTESAL_CTR") & ")"
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If RotinaGerada Then
		bsShowmessage("Não é permitido excluir esta lotação pois a rotina já está gerada!","E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If RotinaGerada Then
		bsShowmessage("Não é permitido incluir lotações pois a rotina já está gerada!","E")
		CanContinue = False
	End If
End Sub
