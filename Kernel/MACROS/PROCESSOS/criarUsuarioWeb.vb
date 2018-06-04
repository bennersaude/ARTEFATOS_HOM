'HASH: F622B0D5FCD7EDC3DB9071B2A2537FCF

Public Sub Main

  Dim viResult, viHPrestador, viHOperadora As Integer 
  Dim vsLogin, vsMensagem As String
  Dim dll As Object

  Set dll = CreateBennerObject("BSPRE001.CriaUsuario")

  viHPrestador = CInt(ServiceVar("HPRESTADOR"))
  viHOperadora = CInt(ServiceVar("HOPERADORA"))
  vsLoginPrestador = CStr(ServiceVar("LOGINPRESTADOR"))

  viResult = dll.Criausuario(CurrentSystem, viHPrestador, vsLoginPrestador, viHOperadora, 0, vsMensagem)

  Set dll = Nothing

  If viResult = 0 Then
    ServiceResult = ""
  Else
    ServiceResult = vsMensagem
  End If

End Sub
