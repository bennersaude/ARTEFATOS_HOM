'HASH: 98ED1BFBE0D9AA78BD54B8BD2507999B
 
Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Dim Q As Object
  Set Q =NewQuery
  Q.Add("Select DATAFINAL FROM SAM_ALERTAMATRICULA  WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_ALERTAMATRICULA")
  Q.Active =True

  If(Not Q.FieldByName("DATAFINAL").IsNull)Then
     MsgBox("O alerta está com a vigência fechada, não pode ser cadastrado mais eventos")
     CanContinue =False
     Exit Sub
  Else
     Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
     Duplica.Duplicar(CurrentSystem,"SAM_ALERTAMATRICULA_EVENTO","ALERTAMATRICULA",RecordHandleOfTable("SAM_ALERTAMATRICULA"),"Duplicando eventos")
     Set Duplica =Nothing
     RefreshNodesWithTable "SAM_ALERTAMATRICULA_EVENTO"
  End If



End Sub
