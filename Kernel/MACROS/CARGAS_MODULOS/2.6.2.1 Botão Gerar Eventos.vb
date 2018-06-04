'HASH: BC1F9D1A4D3091F329C7217649F0D89F
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTAPLANO  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAPLANO")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
     Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
     Duplica.Duplicar(CurrentSystem,"SAM_ALERTAPLANO_EVENTO","PLANOALERTA",RecordHandleOfTable("SAM_ALERTAPLANO"),"Duplicando eventos para alerta")
     Set Duplica =Nothing
     RefreshNodesWithTable "SAM_ALERTAPLANO_EVENTO"
  End If

End Sub
