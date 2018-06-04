'HASH: F6E72B112277A0F94FADCC6D2E6F86FA
 

Public Sub TABLE_AfterScroll()
	Dim qAux As Object
    Set qAux = NewQuery
	qAux.Clear
	qAux.Add("SELECT FRANQUIAINTERNACAO FROM SAM_PARAMETROSPROCCONTAS")
	qAux.Active = True
	If qAux.FieldByName("FRANQUIAINTERNACAO").AsString <> "S" Then
		DIARIAFRANQUIAINTERNACAO.Visible = False
	Else
	    DIARIAFRANQUIAINTERNACAO.Visible = True
	End If
	Set qAux = Nothing
End Sub
