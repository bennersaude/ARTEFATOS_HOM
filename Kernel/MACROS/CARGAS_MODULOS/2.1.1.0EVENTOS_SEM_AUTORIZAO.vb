'HASH: EB1994537A7ABCB8CFD67357D3228B01
Public Sub GERAREVENTOS_OnClick()
Dim Duplica As Object
Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
Duplica.Duplicar(CurrentSystem,"SAM_PLANO_EVENTOSEMAUTORIZ","PLANO",RecordHandleOfTable("SAM_PLANO"),"Duplicando eventos para carência")
Set Duplica =Nothing
RefreshNodesWithTable "SAM_PLANO_EVENTOSEMAUTORIZ"
End Sub
