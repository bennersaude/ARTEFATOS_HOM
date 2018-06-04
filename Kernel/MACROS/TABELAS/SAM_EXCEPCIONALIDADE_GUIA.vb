'HASH: 25E229DA0DA292D36AB66A3A1F363C3F
 

Public Sub ORIGEMGUIA_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False
	
  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="CODIGO|DESCRICAO"
  	vCampos ="Código|Descrição"

	If IsNumeric(ORIGEMGUIA.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SIS_ORIGEMGUIA", vColunas, vNumeroColuna, vCampos, vCriterio, "Origem da guia", False, ORIGEMGUIA.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("ORIGEMGUIA").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
