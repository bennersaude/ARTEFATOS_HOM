'HASH: D43FABD8821497BD7DAB404CE59ECA86


Public Sub AGENDAMENTOS_OnClick()
Dim obj As Object
Set obj =CreateBennerObject("CSCUBEFORMS.CUBOS")
obj.Agendar(CurrentSystem)
Set obj =Nothing
End Sub

Public Sub EXPORTAR_OnClick()
Dim obj As Object
Set obj =CreateBennerObject("CSCUBEFORMS.CUBOS")
obj.Exportar(CurrentSystem)
Set obj =Nothing
End Sub

Public Sub IMPORTAR_OnClick()
Dim obj As Object
Set obj =CreateBennerObject("CSCUBEFORMS.CUBOS")
obj.Importar(CurrentSystem)
Set obj =Nothing
RefreshNodesWithTable("Z_CUBOS")
End Sub
