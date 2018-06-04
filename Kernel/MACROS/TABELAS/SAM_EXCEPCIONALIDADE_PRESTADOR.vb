'HASH: CD6A954B0EF5A04E37619AA34B508396
 

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False

  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="CPFCNPJ|NOME"
  	vCampos ="CPF/CNPJ|Nome"

	If IsNumeric(PRESTADOR.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_PRESTADOR", vColunas, vNumeroColuna, vCampos, vCriterio, "Prestador", True, PRESTADOR.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("PRESTADOR").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
