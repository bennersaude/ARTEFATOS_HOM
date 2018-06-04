'HASH: AEB3F230572CB1EC48F23357E009F680

'MACRO : SAM_BENEFICIARIO_CONTPF

'Public Sub TABLE_AfterScroll()
'	Dim qBuscaRegra
'	Set qBuscaRegra = NewQuery
'	qBuscaRegra.Active = False
'	qBuscaRegra.Add("SELECT DESCRICAO FROM SAM_TABPF WHERE HANDLE = ")
'	If (CurrentQuery.FieldByName("ORIGEMREGRA").AsString = "C") Then
'		qBuscaRegra.Add("(SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pREGRA)")
'	ElseIf (CurrentQuery.FieldByName("ORIGEMREGRA").AsString = "F") Then
'		qBuscaRegra.Add("(SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = (SELECT TABELAPFEVENTO FROM SAM_FAMILIA_PFEVENTO WHERE HANDLE = :pREGRA))")
'	ElseIf (CurrentQuery.FieldByName("ORIGEMREGRA").AsString = "B") Then
'		qBuscaRegra.Add("(SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = (SELECT TABELAPFEVENTO FROM SAM_BENEFICIARIO_PFEVENTO WHERE HANDLE = :pREGRA))")
'	End If
'	qBuscaRegra.ParamByName("pREGRA").AsInteger = CurrentQuery.FieldByName("REGRAPF").AsInteger
'	qBuscaRegra.Active = True
'	If (qBuscaRegra.EOF) Then
'		ROTULOREGRA.Text = "Regra : <Não informada>"
'	Else
'		ROTULOREGRA.Text = "Regra : " + qBuscaRegra.FieldByName("DESCRICAO").AsString
'	End If
'	Set qBuscaRegra = Nothing
'End Sub

