'HASH: 72C55760D6E5B806AD27F45C714D1598
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim qEstado As Object
  Set qEstado =NewQuery
  qEstado.Active =False
  qEstado.Add("SELECT ESTADO FROM SAM_ALERTAESTADO WHERE HANDLE = :HANDLE")
  qEstado.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAESTADO")
  qEstado.Active =True
  If checkPermissao(CurrentSystem,CurrentUser,"E",qEstado.FieldByName("ESTADO").AsInteger,"I")<>"S" Then
    MsgBox "Permissão negada! Usuário não pode executar essa operação."
    Exit Sub
  End If
  
  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTAESTADO  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAESTADO")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
     Dim Duplica As Object
     Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
     Duplica.Duplicar(CurrentSystem,"SAM_ALERTAESTADO_EVENTO","ALERTAESTADO",RecordHandleOfTable("SAM_ALERTAESTADO"),"Duplicando eventos para alerta")
     Set Duplica =Nothing
     RefreshNodesWithTable "SAM_ALERTAESTADO_EVENTO"
  End If

End Sub
