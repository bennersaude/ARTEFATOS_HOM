'HASH: 6F7BCC761F6A63B12DB8CBBC6D89A502
Option Explicit
 '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
If CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime <> 0 Then
  If CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime < CurrentQuery.FieldByName("COMPETENCIA").AsDateTime Then
    bsShowMessage("Competência final menor que a competência inicial","I")
    CanContinue = False
  End If

  If CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime And CurrentQuery.FieldByName("COMPETENCIA").AsDateTime = 0 Then
    bsShowMessage("Competência inicial deve ser preenchida","I")
    CanContinue = False
  End If
End If
End Sub
