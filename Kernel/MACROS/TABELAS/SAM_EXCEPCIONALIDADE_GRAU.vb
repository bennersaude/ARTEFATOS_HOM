'HASH: EFFA09BD6EF1618C852E098EDB886423


Public Sub GRAU_OnPopup(ShowPopup As Boolean)
 	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False

  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="GRAU|DESCRICAO"
  	vCampos ="Grau|Descrição"

	If IsNumeric(GRAU.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_GRAU", vColunas, vNumeroColuna, vCampos, vCriterio, "Grau", True, GRAU.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("GRAU").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
