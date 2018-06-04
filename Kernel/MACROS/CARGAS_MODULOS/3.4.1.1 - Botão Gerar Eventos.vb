'HASH: CB49C96177B629D4C4FE2AD34C84FE56
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  
  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTAGERAL  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAGERAL")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
     Dim Duplica As Object
     Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
     Duplica.Duplicar(CurrentSystem,"SAM_ALERTAGERAL_EVENTO","ALERTAGERAL",RecordHandleOfTable("SAM_ALERTAGERAL"),"Duplicando eventos para alerta")
     Set Duplica =Nothing
     RefreshNodesWithTable "SAM_ALERTAGERAL_EVENTO"
  End If

End Sub
