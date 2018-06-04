'HASH: AA18ACA79932E5B7FFCD83B9ED8ED143
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim SQL As Object

	Set SQL = NewQuery
	SQL.Add("SELECT HANDLE FROM SIS_TIPOFATURAMENTO WHERE CODIGO=110 AND HANDLE=" + CStr(CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger))
	SQL.Active =True

    If Not SQL.EOF Then
    	bsShowMessage("Existe outro relatório para listar os Acertos de Custo Operacional !" + Chr(13) + _
    				  "DEM-AC5.DEMONSTRATIVO DE FATURAMENTO CUSTO OPERACIONAL - ACERTO", "E")
    	CanContinue = False
    	Exit Sub
    End If
End Sub
