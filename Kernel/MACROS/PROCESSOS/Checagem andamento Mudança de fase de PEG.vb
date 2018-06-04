'HASH: 7724CA7B36D46709FFFEF429FF74B88C

Public Sub Main

  Dim qAux As Object
  Set qAux = NewQuery

  qAux.Clear
  qAux.Add("UPDATE SAM_PEG SET SITUACAOPROCESSAMENTO = 1 WHERE SITUACAOPROCESSAMENTO IN (2, 4) AND DATA <= :DATA")
  qAux.ParamByName("DATA").AsDateTime = ServerNow - 0.25
  qAux.ExecSQL

  Set qAux = Nothing

End Sub
