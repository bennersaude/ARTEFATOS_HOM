'HASH: 05AF57A042D7D96B0910CD761A2AC333
 

Public Sub REGIMEATENDIMENTO_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
 	Dim vHandle As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vNumeroColuna As Integer

	ShowPopup = False
	
  	Set interface =CreateBennerObject("Procura.Procurar")

  	vColunas ="CODIGO|DESCRICAO"
  	vCampos ="Código|Descrição"

	If IsNumeric(REGIMEATENDIMENTO.Text) Then
		vNumeroColuna = 1

	Else
		vNumeroColuna = 2

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_REGIMEATENDIMENTO", vColunas, vNumeroColuna, vCampos, vCriterio, "Regime de atendimento", True, REGIMEATENDIMENTO.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("REGIMEATENDIMENTO").Value =vHandle
  	End If
  	Set interface =Nothing
End Sub
