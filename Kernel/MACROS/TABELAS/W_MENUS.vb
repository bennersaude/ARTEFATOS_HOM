'HASH: 7669D3C3390263C56B072DB0E414DD0F
 
Public Sub NOME_OnExit() 
  If CurrentQuery.FieldByName("CODIGO").AsString = "" And CurrentQuery.State = 3 Then 
    CurrentQuery.FieldByName("CODIGO").AsString = CurrentQuery.FieldByName("NOME").AsString 
  End If 
End Sub 
