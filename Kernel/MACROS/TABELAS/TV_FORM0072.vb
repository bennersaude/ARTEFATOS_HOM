'HASH: 7454D798C1C352BD2B75DD7ECCD3DBEA
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Interface  As Object
  Dim viRetorno  As Integer
  Dim vsMensagem As String

  Set Interface = CreateBennerObject("BSDMED.Rotinas")
  viRetorno = Interface.GerarArquivo(CurrentSystem, CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger, vsMensagem)

  If viRetorno > 0 Then
    CanContinue = False
    bsShowMessage(vsMensagem, "E")
    Exit Sub
  Else
    bsShowMessage("Processo enviado para execução no servidor!", "I")
  End If

End Sub
