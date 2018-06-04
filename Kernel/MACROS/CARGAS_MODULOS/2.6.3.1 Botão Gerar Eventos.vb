'HASH: 3D879FACD7400618F28BE88096F51F49
Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object

  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTACONTRATO  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTACONTRATO")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
Duplica.Duplicar(CurrentSystem,"SAM_ALERTACONTRATO_EVENTO","CONTRATOALERTA",RecordHandleOfTable("SAM_ALERTACONTRATO"),"Duplicando eventos para alerta")
Set Duplica =Nothing
RefreshNodesWithTable "SAM_ALERTACONTRATO_EVENTO"
  End If
End Sub
