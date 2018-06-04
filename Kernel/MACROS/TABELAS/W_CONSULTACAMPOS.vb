'HASH: D79682039481FD20DE6AD72A3421A118

Public Sub TABLE_AfterPost()
  x = RecordHandleOfTable("W_CONSULTAS")
  Set Q = NewQuery
  Q.Add("UPDATE W_CONSULTAS SET ULTIMAALTERACAO = :DATA WHERE HANDLE = " + CStr(x))
  Q.ParamByName("DATA").AsDateTime = Now
  Q.ExecSQL
  Set Q = Nothing
End Sub

