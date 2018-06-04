'HASH: 1BE00CCAF7B2C9F19C86724EA63048A5
 
Option Explicit
'#Uses "*bsShowMessage"

Public Sub SUSCALCULARHASH_OnClick()

Dim obj As Object
	Set obj = CreateBennerObject("BENNER.SAUDE.SERVICES.PROCCONTAS.RESSARCIMENTOSUS.Rotinas")

	Dim vHash As String
	Dim vArquivo As String
	Dim vResultado As String

	vArquivo = OpenDialog

	If vArquivo <> "" Then
		vResultado = obj.CalcularHash(CurrentSystem,vArquivo)
		If VisibleMode Then
    		bsShowMessage("O Hash do arquivo informado é: " + vResultado, "I")
    	Else
			CancelDescription = "O Hash do arquivo informado é: " + vResultado
		End If
	End If
	Set obj = Nothing


End Sub

Public Sub SUSCARREGARXML_OnClick()

	Dim sx As Object

	Set sx = CreateBennerObject("BENNER.SAUDE.SERVICES.PROCCONTAS.RESSARCIMENTOSUS.ImportacaoRessarcimentoSUS")

	sx.VarrerDiretorioXML(CurrentSystem)

	If VisibleMode Then
		bsShowMessage("Processo concluido com sucesso!", "I")
		RefreshNodesWithTable("SAM_ROTRESSARCIMENTOSUS")
	Else
    	CancelDescription = "Processo concluido com sucesso!"
	End If

	Set sx = Nothing

End Sub
