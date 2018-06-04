'HASH: 94B0467DCD9170C42C98E8AB4A859F26
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT COUNT(HANDLE) NREC FROM SAM_PROCREF_PRESTADOR WHERE PROCREF = :PROCREF")
  q1.ParamByName("PROCREF").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True
  If q1.FieldByName("NREC").AsInteger > 0 Then
    bsShowMessage("Este processo já possui prestadores cadastrados. Exclusão cancelada !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub




Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT HANDLE FROM SAM_PROCREF_AVALIACAOREF WHERE HANDLE <> :HANDLE AND AVALIACAOREF = :AVALIACAOREF AND PROCREF = :PROCREF")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  q1.ParamByName("AVALIACAOREF").Value = CurrentQuery.FieldByName("AVALIACAOREF").Value
  q1.ParamByName("PROCREF").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True
  If Not q1.EOF Then
    bsShowMessage("Esta avaliação já está cadastrada neste processo!", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

