'HASH: 90A758CEE681499476F394305477F6D2
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*bsShowMessage"

Option Explicit
Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

 	Dim interface As Object
 	Dim vHandle As Long
 	Dim vCampos As String
  	Dim vColunas As String
  	Dim vCriterio As String
  	Dim vNumeroColuna As Integer

  	ShowPopup = False

 	Set interface =CreateBennerObject("Procura.Procurar")

	vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
  	vColunas ="ESTRUTURA|ESTRUTURANUMERICA|DESCRICAOABREVIADA|NIVELAUTORIZACAO"
  	vCampos ="Estrutura|Estrutura Numérica|Descrição|Nível de autorização"

	If IsNumeric(EVENTO.Text) Then
		vNumeroColuna = 2

	Else
		vNumeroColuna = 3

	End If

    vHandle =interface.Exec(CurrentSystem,"SAM_TGE",vColunas, vNumeroColuna, vCampos, vCriterio, "Evento", False, EVENTO.Text)

  	If vHandle <>0 Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("EVENTO").Value = vHandle
    	CurrentQuery.FieldByName("ESTRUTURA").Value = vHandle
  	End If

  	Set interface =Nothing
End Sub
