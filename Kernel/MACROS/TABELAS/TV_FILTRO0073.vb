'HASH: 5A36E45379F302D81F8685A6FC046332
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLFiltro As Object

     Set SqltipoFat =NewQuery
    SqltipoFat.Clear
    SqltipoFat.Add("SELECT HANDLE, CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE=" + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
    SqltipoFat.Active =True

    If((SqltipoFat.FieldByName("CODIGO").AsInteger)<>630)Then
	    bsShowMessage("O TIPO DE FATURAMENTO deve ser recolhimento de IRRF !", "E")
	    CanContinue = False
	    Exit Sub
	End If
End Sub
