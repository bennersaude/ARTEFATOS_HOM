'HASH: 60A048D9B333FA28EC515BAE38DB5C79
 

Public Sub BOTAOGERAREVENTOS_OnClick()
Dim qMunicipio As Object
Set qMunicipio =NewQuery
qMunicipio.Active =False
qMunicipio.Add("SELECT MUNICIPIO FROM SAM_ALERTAMUNICIPIO WHERE HANDLE = :HANDLE")
qMunicipio.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAMUNICIPIO")
qMunicipio.Active =True
If checkPermissao(CurrentSystem,CurrentUser,"M",qMunicipio.FieldByName("MUNICIPIO").AsInteger,"I")<>"S" Then
    MsgBox "Permissão negada! Usuário não pode executar essa operação."
    Set qMunicipio =Nothing
    Exit Sub
  End If

  Dim Duplica As Object
  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTAMUNICIPIO  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAMUNICIPIO")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_ALERTAMUNICIPIO_EVENTO","ALERTAMUNICIPIO",RecordHandleOfTable("SAM_ALERTAMUNICIPIO"),"Duplicando eventos para alerta")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_ALERTAMUNICIPIO_EVENTO"
  End If


End Sub
