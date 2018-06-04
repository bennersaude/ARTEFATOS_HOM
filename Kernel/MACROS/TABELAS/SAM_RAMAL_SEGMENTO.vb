'HASH: AEDB26E6CCDD6147171CC86780B73B9E
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim qConsulta As Object
	Set qConsulta = NewQuery

	qConsulta.Active = False
	qConsulta.Clear
	qConsulta.Add("   SELECT HANDLE                                      ")
	qConsulta.Add("     FROM SAM_RAMAL_SEGMENTO                          ")
	qConsulta.Add("    WHERE RAMAL = :PRAMAL                             ")
	qConsulta.Add("      AND SEGMENTO = :PSEGMENTO ")
	qConsulta.Add("      AND HANDLE <> :PHANDLE                          ")
	qConsulta.ParamByName("PRAMAL").AsInteger    = CurrentQuery.FieldByName("RAMAL").AsInteger
	qConsulta.ParamByName("PSEGMENTO").AsInteger = CurrentQuery.FieldByName("SEGMENTO").AsInteger
	qConsulta.ParamByName("PHANDLE").AsInteger   = CurrentQuery.FieldByName("HANDLE").AsInteger
	qConsulta.Active = True

	If Not qConsulta.EOF Then
		bsShowMessage("Já existe um registro salvo com esse mesmo segmento.", "E")
		CanContinue = False
	End If

	Set qConsulta = Nothing

End Sub
