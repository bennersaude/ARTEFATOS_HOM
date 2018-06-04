'HASH: 1551E5ED44BB25EA1287DF1CAA79235D
Option Explicit

Public Sub GERAREVENTOS_OnClick()
	Dim Duplica As Object

	Set Duplica=CreateBennerObject("SamDupEventos.Rotinas")

	Duplica.Duplicar(CurrentSystem, "SAM_PREST_TXCOMERCIALIZACAO_EV","TAXACOMERCIALIZACAO",RecordHandleOfTable("SAM_PREST_TXCOMERCIALIZACAO"), "Buscar evento", 0, "")

	Set Duplica=Nothing

	RefreshNodesWithTable"SAM_PREST_TXCOMERCIALIZACAO_EV"
End Sub
