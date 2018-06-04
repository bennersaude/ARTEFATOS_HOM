'HASH: CC936AB9C4F2F766D3DF0F6EF552CC46
 

Public Sub IMPORTALOTE_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SAMPEGDIGIT.DIGITACAO")
    interface.IMPORTARPEGLOTE(CurrentSystem)
  Set interface =Nothing
End Sub
