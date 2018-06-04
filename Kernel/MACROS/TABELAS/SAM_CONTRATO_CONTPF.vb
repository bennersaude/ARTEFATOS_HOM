'HASH: 537F3B5D013B560064DB049B7E7E5360

'MACRO : SAM_CONTRATO_CONTPF

'Public Sub TABLE_AfterScroll()
'	Dim qBuscaRegra
'	Set qBuscaRegra = NewQuery
'	qBuscaRegra.Active = False
'	qBuscaRegra.Add("SELECT DESCRICAO FROM SAM_TABPF WHERE HANDLE = ")
'	qBuscaRegra.Add("(SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pREGRA)")
'	qBuscaRegra.ParamByName("pREGRA").AsInteger = CurrentQuery.FieldByName("REGRAPF").AsInteger
'	qBuscaRegra.Active = True
'	If (qBuscaRegra.EOF) Then
'		ROTULOREGRA.Text = "Regra : <Não informada>"
'	Else
'		ROTULOREGRA.Text = "Regra : " + qBuscaRegra.FieldByName("DESCRICAO").AsString
'	End If
'	Set qBuscaRegra = Nothing
'End Sub

