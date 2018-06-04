'HASH: C8DCA2D5D4F2C6EE1074F240E809430D

Public Sub Main
	Dim DLL As Object

	On Error GoTo Erro
	Set DLL = CreateBennerObject("ROTARQ.ROTINAS")
	DLL.CriarProcessarRetornoCompras(CurrentSystem, SessionVar("DIRETORIO_IMPORTACAO"), SessionVar("DIRETORIO_OK"), SessionVar("DIRETORIO_ERRO"))
	Set DLL = Nothing
	Exit Sub

	Erro:
	Set DLL = Nothing
	Err.Raise(Err.Number, Err.Source, Err.Description)

End Sub
