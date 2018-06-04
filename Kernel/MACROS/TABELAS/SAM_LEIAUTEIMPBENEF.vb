'HASH: 52054DB8A8C354FAA4F61A55CE81692A

Public Sub TABLE_AfterInsert()
    If (VisibleMode And NodeInternalCode = 1880) Or (WebMode And WebMenuCode = "T1880") Then
  		CurrentQuery.FieldByName("LEIAUTEDE").AsInteger = 2
  	Else
  		CurrentQuery.FieldByName("LEIAUTEDE").AsInteger = 1
	End If
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime = ServerDate
  CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = CurrentUser
End Sub

Public Sub TABLE_UpdateRequired()
  If CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = 0 Then
    CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime = ServerDate
    CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = CurrentUser
  End If
End Sub
