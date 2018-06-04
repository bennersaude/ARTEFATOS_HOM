'HASH: 5FDB5019A6862748B2D0252554214559
Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Dim Q As Object

  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTAFAMILIA  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAFAMILIA")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
     Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
     Duplica.Duplicar(CurrentSystem,"SAM_ALERTAFAMILIA_EVENTO","FAMILIAALERTA",RecordHandleOfTable("SAM_ALERTAFAMILIA"),"Duplicando eventos para alerta")
     Set Duplica =Nothing
     RefreshNodesWithTable "SAM_ALERTAFAMILIA_EVENTO"
  End If

End Sub
