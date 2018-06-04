'HASH: A1D86AD202CCB95BC58697382AC5887E
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem,"I","P",Msg)="N" Then
    MsgBox Msg
    CanContinue =False
    Exit Sub
  End If

  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_PRESTADOR_BONIFICA_EVENTO","PRESTADORBONIFICA",RecordHandleOfTable("SAM_PRESTADOR_BONIFICA"),"Gerando eventos")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_PRESTADOR_BONIFICA_EVENTO"
End Sub
