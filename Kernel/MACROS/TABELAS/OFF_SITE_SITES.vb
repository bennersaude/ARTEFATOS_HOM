'HASH: 39F8F99F67342C735DF1DA73CDEF76B5
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Active = False
Consulta.Add("SELECT OFF_SITE.HANDLE                                          ")
Consulta.Add("  FROM OFF_SITELOCAL OFL                                        ")
Consulta.Add("  JOIN OFF_SITE      OFF_SITE  ON OFF_SITE.HANDLE = OFL.NOMESITE")
Consulta.Add(" WHERE OFL.HANDLE = :HANDLE                                     ")
Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("SITELOCAL").AsInteger
Consulta.Active = True


If CurrentQuery.FieldByName("NOMESITE").Value = Consulta.FieldByName("HANDLE").Value Then
  bsShowMessage("Site a replicar deve ser diferente do site local!", "I")
  CanContinue = False
  Exit Sub
End If


End Sub
