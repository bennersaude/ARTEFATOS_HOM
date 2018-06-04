'HASH: 4BAAA761BEC67EA768683693F219A7DD
 

Public Sub TIPOPRESTADOR_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False
	
  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="CODIGO|DESCRICAO"
  	vCampos ="Código|Descrição"

	If IsNumeric(TIPOPRESTADOR.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_TIPOPRESTADOR", vColunas, vNumeroColuna, vCampos, vCriterio, "Tipo de prestador", True, TIPOPRESTADOR.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("TIPOPRESTADOR").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
