'HASH: CB1B0E051C06B5A225646244E9280BDF

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object

  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTABENEF  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTABENEF")
  Q.Active =True


  If(Not Q.FieldByName("DATAFINAL").IsNull) And (ServerDate > Q.FieldByName("DATAFINAL").AsDateTime) Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
     Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
     Duplica.Duplicar(CurrentSystem,"SAM_ALERTABENEF_EVENTO","BENEFICIARIOALERTA",RecordHandleOfTable("SAM_ALERTABENEF"),"Duplicando eventos")
     Set Duplica =Nothing
     RefreshNodesWithTable "SAM_ALERTABENEF_EVENTO"
  End If

End Sub
