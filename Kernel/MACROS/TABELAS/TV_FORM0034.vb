'HASH: 0ED8199D685D74679F830B45EEBD95AD
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim dllBSBen022 As Object
  Dim viResult    As Integer
  Dim vsMensagem  As String

  Set dllBSBen022 = CreateBennerObject("BSBEN022.Modulo")

  viResult = dllBSBen022.Reativar(CurrentSystem, _
                                  CLng(SessionVar("HMODBENEFICIARIO")), _
                                  CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime, _
                                  vsMensagem)

  If viResult = 1 Then
    bsShowMessage(vsMensagem, "E")
    CanContinue = False
  Else
    bsShowMessage("Reativação concluída!", "I")
  End If

  Set dllBSBen022 = Nothing
End Sub
