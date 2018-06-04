'HASH: CD469CB6AE832CD04C5BD9FCFE089BCD
 

Public Sub NIVELAUTORIZ_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False
	
  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="CODIGO|DESCRICAO"
  	vCampos ="Código|Descrição"

	If IsNumeric(NIVELAUTORIZ.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_NIVELAUTORIZACAO", vColunas, vNumeroColuna, vCampos, vCriterio, "Nível de autorização", True, NIVELAUTORIZ.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("NIVELAUTORIZ").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
