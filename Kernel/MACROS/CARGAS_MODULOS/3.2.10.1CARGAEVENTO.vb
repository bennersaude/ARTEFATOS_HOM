'HASH: 3ECF6075F486A08C30BBE6932141FCD0
Option Explicit

Public Sub INCLUIR_OnClick()
	Dim Duplica As Object

	Set Duplica=CreateBennerObject("SamDupEventos.Rotinas")

	Duplica.Duplicar(CurrentSystem, "SAM_EXCEPCIONALIDADE_EVENTO","EXCEPCIONALIDADE",RecordHandleOfTable("SAM_EXCEPCIONALIDADE"), "Buscar evento", 0, "")

	Set Duplica=Nothing

	RefreshNodesWithTable"SAM_EXCEPCIONALIDADE_EVENTO"
End Sub
