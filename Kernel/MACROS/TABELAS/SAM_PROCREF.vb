'HASH: 12732D4EE51FAB76C736B0A05E7CB035


Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser
End Sub

