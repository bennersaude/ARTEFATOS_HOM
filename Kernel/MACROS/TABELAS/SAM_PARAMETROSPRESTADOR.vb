'HASH: 12444567405A067CA864B17FEB77FB83
'MACRO = SAM_PARAMETROSPRESTADOR
'#Uses "*bsShowMessage"

Option Explicit


Public Sub TABLE_AfterScroll()
  TextAjuda(CurrentQuery.FieldByName("TABPADRAOCODIGO").AsInteger)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If (TABPADRAOCODIGO.PageIndex = 1) Or (TABPADRAOCODIGO.PageIndex = 2) Then
    If (CurrentQuery.FieldByName("PADRAOCONSELHO").AsInteger < 4) Then
      CanContinue = False
      BsShowMessage("Inscrição do conselho obrigatória.", "I")
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABPADRAOCODIGO").AsInteger <= 2 Then
    If CurrentQuery.FieldByName("SOBREPORALTERACAO").AsString = "N" And CurrentQuery.FieldByName("DIGITARPRESTADOR").AsString = "N" Then
      CanContinue = False
      BsShowMessage( "Caso não sobreponha as alterações, deve ser digitado o prestador.", "I")
      Exit Sub
    End If
  End If

  If (CurrentQuery.FieldByName("TABPADRAOCODIGO").AsInteger = 3) Or (CurrentQuery.FieldByName("TABPADRAOCODIGO").AsInteger = 3) Then
    If InStr(CurrentQuery.FieldByName("MASCARAPRESTADOR").AsString, "9") = 0 Then
      CanContinue = False
      BsShowMessage( "Máscara obrigatória.", "I")
      Exit Sub
    End If
  End If

  Dim cont As Integer
  Dim i As Integer

  cont = 0
  For i = 1 To Len(CurrentQuery.FieldByName("MASCARAPRESTADOR").AsString)
    If Mid(CurrentQuery.FieldByName("MASCARAPRESTADOR").AsString, i, 1) = "9" Then cont = cont + 1
  Next

  If cont > 9 Then
    CanContinue = False
    BsShowMessage( "Máscara deve conter no máximo 9 dígitos.", "I")
    Exit Sub
  End If


End Sub

Public Sub TABPADRAOCODIGO_OnChange()
  TextAjuda(TABPADRAOCODIGO.PageIndex)
End Sub

Public Sub TextAjuda(pTabPadraoCodigo As Byte)
  Select Case pTabPadraoCodigo
    Case 0
      TEXTOAJUDA.Text = "Caso não seja informado o CPF/CNPJ será utilizado o código do conselho como chave."
    Case 1
      TEXTOAJUDA.Text = "Caso não seja informado o código do conselho será utilizado o CPF/CNPJ como chave."
    Case 2
      TEXTOAJUDA.Text = "Chave gerada pelo sistema."
    Case 3
      TEXTOAJUDA.Text = "Chave informada pelo sistema, conforme a máscara."
  End Select
End Sub
