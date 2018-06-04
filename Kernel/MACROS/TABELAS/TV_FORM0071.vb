'HASH: 283697AD61ECFC8EFCF244DB52DA590A

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Interface  As Object
  Dim viRetorno  As Integer
  Dim vsMensagem As String

  Set Interface = CreateBennerObject("BSDMED.Rotinas")
  viRetorno = Interface.Retificar(CurrentSystem, CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger, vsMensagem)

  If viRetorno > 0 Then
    CanContinue = False
    bsShowMessage(vsMensagem, "E")
    Exit Sub
  Else
    bsShowMessage("Processo enviado para execução no servidor!", "I")
  End If

End Sub

