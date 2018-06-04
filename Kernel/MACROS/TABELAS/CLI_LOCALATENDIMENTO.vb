'HASH: 453A5A65AE4BB2B8E4541E8FA182DD3D
Option Explicit

'#uses "*bsShowMessage

Public Sub TABLE_AfterPost()
 If (WebMode) Then
   If CurrentQuery.FieldByName("ASSUMIRPADRAO").AsString = "S" Then
     Dim SQL As Object
     Set SQL = NewQuery
     SQL.Active = False
     SQL.Clear
     SQL.Add("UPDATE CLI_LOCALATENDIMENTO ")
     SQL.Add("   SET ASSUMIRPADRAO = 'N'  ")
     SQL.Add(" WHERE HANDLE <> :HANDLE    ")
     SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
     SQL.ExecSQL
     bsShowMessage("Este local de atendimento será utilizado como padrão para o SOAP","I")
   End If
 End If
End Sub
