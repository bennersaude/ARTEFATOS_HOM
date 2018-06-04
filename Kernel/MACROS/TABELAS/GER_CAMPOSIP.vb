'HASH: 802C53B0DBAAE5AC3262E879DECC59B1
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim q As Object
  Set q = NewQuery
  q.Clear

  q.Add("SELECT 1                    ")
  q.Add("  FROM GER_CAMPOSIP         ")
  q.Add(" WHERE CAMPOSIP = :CAMPOSIP ")
  q.Add("   AND HANDLE <> :HANDLE    ")

  q.ParamByName("CAMPOSIP").Value = CurrentQuery.FieldByName("CAMPOSIP").Value
  q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  q.Active = True

  If Not q.EOF Then
    CanContinue = False
   	MsgBox("Este Campo SIP já está cadastrado !")

  End If

End Sub
