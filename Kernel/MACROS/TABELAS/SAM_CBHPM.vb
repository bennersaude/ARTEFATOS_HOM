'HASH: A6257CDE10750D8A603E34A6C141FB5F

Public Sub TABLE_AfterScroll()
  CurrentQuery.FieldByName("ESTRUTURA").Mask = "9.99.99.999-9;9"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not VisibleMode = True Then
    Exit Sub
  End If

  Dim SQL As Object
  Dim SQLNIV As Object
  Dim QueryNiveis As Object
  Dim vEstruturaSemMascara As String
  Dim vQtdeNumerosMascara As Integer



  ' Tratamento da Estrutura da SAM_TGE

  vEstruturaSemMascara = ""
  vQtdeNumerosMascara = 0

  For vIndice = 1 To Len(CurrentQuery.FieldByName("ESTRUTURA").AsString)
    If InStr("0123456789", Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, vIndice, 1)) > 0 Then
      vEstruturaSemMascara = vEstruturaSemMascara + _
                             Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, vIndice, 1)
    End If
  Next vIndice


  CurrentQuery.FieldByName("ESTRUTURANUMERICA").Value = Val(vEstruturaSemMascara)

End Sub



Public Sub AtualizaEstruturaNumerica

  Dim sValor As String
  Dim i As Long

  sValor = ""

  For i = 1 To Len(CurrentQuery.FieldByName("ESTRUTURA").AsString)
    If InStr("0123456789", Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, i, 1)) > 0 Then
      sValor = sValor + Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, i, 1)
    End If
  Next i

  CurrentQuery.FieldByName("ESTRUTURANUMERICA").Value = Val(sValor)

End Sub

