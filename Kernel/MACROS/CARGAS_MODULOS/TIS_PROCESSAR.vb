'HASH: 144710F55A777167DEB3B94C64859ED5
Option Explicit

Public Sub TISHASH_OnClick()
	Dim obj As Object
	Set obj = CreateBennerObject("BSTISS.Rotinas")

	Dim vHash As String
	Dim vArquivo As String
	Dim vResultado As String

	vArquivo = OpenDialog

	If vArquivo <> "" Then
		vResultado = obj.Hash(CurrentSystem,vArquivo)
    	MsgBox(vResultado)
	End If


	Set obj = Nothing
 
End Sub

Public Sub TISPROCESSAR_OnClick()
	Dim Interface As Object
	Set Interface = CreateBennerObject("BSTISS.ROTINAS")
	Interface.CarregarXML(CurrentSystem)
	Set Interface = Nothing
End Sub
