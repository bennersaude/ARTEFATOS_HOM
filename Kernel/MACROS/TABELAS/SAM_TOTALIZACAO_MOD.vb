'HASH: F7E2A4749FA9CF91D3463E256FAAC25F
 
Public Sub TABLE_AfterScroll()
	Dim SQL As Object

	Set SQL = NewQuery
	SQL.Add("SELECT ABERTURAPLANOFAIXA, ABERTURACONTRATOPROD FROM SAM_MODELO")
	SQL.Active = True

    FAIXAETARIAINICIAL.ReadOnly = False
	PLANO.ReadOnly = False
	MODULO.ReadOnly = False
    CONTRATO.ReadOnly  = False
    TIPOPRODUTO.ReadOnly = False

	If SQL.FieldByName("ABERTURACONTRATOPROD").AsString = "N" Then
	   CONTRATO.ReadOnly = True
	   TIPOPRODUTO.ReadOnly = True
	End If


	If SQL.FieldByName("ABERTURAPLANOFAIXA").AsString = "N" Then
	   FAIXAETARIAINICIAL.ReadOnly = True
	   MODULO.ReadOnly = True
	   PLANO.ReadOnly = True
	End If
    Set SQL = Nothing
End Sub
