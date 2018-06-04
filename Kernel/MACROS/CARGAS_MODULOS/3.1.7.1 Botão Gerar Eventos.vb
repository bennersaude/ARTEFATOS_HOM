'HASH: 635132D60F540DA6AD9E46279E5D67BD
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem,"I","P",Msg)="N" Then
    MsgBox Msg
    CanContinue =False
    Exit Sub
  End If
  
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_PRESTADOR_MAJORA_EVENTO","PRESTADORMAJORA",RecordHandleOfTable("SAM_PRESTADOR_MAJORA"),"Gerando eventos")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_PRESTADOR_MAJORA_EVENTO"
End Sub
