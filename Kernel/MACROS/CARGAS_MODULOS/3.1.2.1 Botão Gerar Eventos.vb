'HASH: 407B9F2222536DF4F233E634BFAE5D0C
 
Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem,"A","P",Msg)="N" Then
    MsgBox Msg
    Exit Sub
  End If

  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_ALERTAPRESTADOR_EVENTO","ALERTAPRESTADOR",RecordHandleOfTable("SAM_ALERTAPRESTADOR"),"Duplicando eventos para alerta")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_ALERTAPRESTADOR_EVENTO"
End Sub
