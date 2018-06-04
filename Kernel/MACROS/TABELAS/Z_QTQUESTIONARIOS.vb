'HASH: DA9D07DE8D4A48C9A7A7A3FC970D48C7
Public Function LimpaStr(Value As String) 
  Caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_" 
  LimpaStr = "" 
  For i = 1 To Len(Value) 
    If InStr(Caracteres, Mid(Value, i, 1)) Then 
      LimpaStr = LimpaStr + Mid(Value, i, 1) 
    End If 
  Next i 
End Function 
 
Public Sub CODIGO_OnExit() 
  If (CurrentQuery.State <> 1) Then 
    CurrentQuery.FieldByName("CODIGO").AsString = LimpaStr(CurrentQuery.FieldByName("CODIGO").AsString) 
  End If 
End Sub 
 
Public Sub TABLE_UpdateRequired() 
  CurrentQuery.FieldByName("CODIGO").AsString = LimpaStr(CurrentQuery.FieldByName("CODIGO").AsString) 
End Sub 
