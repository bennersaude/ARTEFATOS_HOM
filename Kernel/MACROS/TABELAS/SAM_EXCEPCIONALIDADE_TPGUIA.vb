'HASH: B6C673DAF6AE739A0FE9728CF157A902
 

Public Sub TIPOGUIA_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False
	
  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="CODIGO|DESCRICAO"
  	vCampos ="Código|Descrição"

	If IsNumeric(TIPOGUIA.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_TIPOGUIA", vColunas, vNumeroColuna, vCampos, vCriterio, "Tipo de guia", True, TIPOGUIA.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("TIPOGUIA").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
