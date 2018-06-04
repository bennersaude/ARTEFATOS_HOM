'HASH: 87BE452B5E34C46211F03133BAAB395E

Public Sub TIPOORIGEM_OnChange()
  If CurrentQuery.FieldByName("TIPOORIGEM").Value = "2" Then
    TIPOCLASSEGERENCIAL.ReadOnly = True
  Else
    TIPOCLASSEGERENCIAL.ReadOnly = False
  End If
  CurrentQuery.UpdateRecord
End Sub

