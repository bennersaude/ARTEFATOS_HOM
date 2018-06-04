'HASH: 4B6618EB45B270AA061E82316C801DBC
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
  q1.Add("SELECT HANDLE FROM SAM_PROCREF_ESPECIALIDADE WHERE HANDLE <> :HANDLE AND ESPECIALIDADE = :ESPECIALIDADE AND PROCREF = :PROCREF")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  q1.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").Value
  q1.ParamByName("PROCREF").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True
  If Not q1.EOF Then
    bsShowMessage("Esta especialidade já está cadastrada neste processo!", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

