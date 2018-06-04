'HASH: 8D989AEB17F64A6A2755ECFF579CC58F
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem,"A","P",Msg)="N" Then
    MsgBox Msg
    Exit Sub
  End If

  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTAPRESTADOR  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAPRESTADOR")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
     Dim Duplica As Object
     Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
     Duplica.Duplicar(CurrentSystem,"SAM_ALERTAPRESTADOR_EVENTO","ALERTAPRESTADOR",RecordHandleOfTable("SAM_ALERTAPRESTADOR"),"Duplicando eventos para alerta")
     Set Duplica =Nothing
     RefreshNodesWithTable "SAM_ALERTAPRESTADOR_EVENTO"
  End If
End Sub
