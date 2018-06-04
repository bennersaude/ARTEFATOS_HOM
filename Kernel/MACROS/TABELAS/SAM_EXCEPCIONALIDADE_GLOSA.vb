'HASH: 5159D2F7E54C03521B8AA2E96C0D21E1
 

Public Sub MOTIVOGLOSA_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False
	
  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="CODIGOGLOSA|DESCRICAO"
  	vCampos ="Código Glosa|Descrição"

	If IsNumeric(MOTIVOGLOSA.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_MOTIVOGLOSA", vColunas, vNumeroColuna, vCampos, vCriterio, "Motivo de glosa", True, MOTIVOGLOSA.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("MOTIVOGLOSA").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
