'HASH: 6B23C59C59DF59C7075D7B0FD9BFE201
Public Sub BOTOGERAEVENTO_OnClick()
Dim Duplica As Object
Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
Duplica.Duplicar(CurrentSystem,"SAM_TIPOAUTORIZ_EVENTO","TIPOAUTORIZ",RecordHandleOfTable("SAM_TIPOAUTORIZ"),"Gerando eventos para o tipo de autorização")
Set Duplica =Nothing
RefreshNodesWithTable "SAM_TIPOAUTORIZ_EVENTO"
End Sub
