'HASH: 6E3FB7D460B1A28371461C664F6C9D1B

Public Sub TABLE_AfterCancel()
  CALOR.Visible = True
  ILUMINANCIA.Visible = True
  RUIDO.Visible = True
  MEDIDA.Caption = "Medida Encontrada"
  MEDIDA.Visible = True
  COMPONENTE.Visible = True
  TIPOILUMINACAO.Visible = True
End Sub

Public Sub TABLE_AfterPost()
  CALOR.Visible = True
  ILUMINANCIA.Visible = True
  RUIDO.Visible = True
  MEDIDA.Caption = "Medida Encontrada"
  MEDIDA.Visible = True
  COMPONENTE.Visible = True
  TIPOILUMINACAO.Visible = True
End Sub

Public Sub TIPO_OnChange()
  'Geral
  If (CurrentQuery.FieldByName("TIPO").AsInteger = 1) Then
    CALOR.Visible = False
    ILUMINANCIA.Visible = False
    RUIDO.Visible = False
    TIPOILUMINACAO.Visible = False
  End If
  'Calor
  If (CurrentQuery.FieldByName("TIPO").AsInteger = 2) Then
    ILUMINANCIA.Visible = False
    RUIDO.Visible = False
    TIPOILUMINACAO.Visible = False
  End If
  'Iluminancia
  If (CurrentQuery.FieldByName("TIPO").AsInteger = 3) Then
    CALOR.Visible = False
    RUIDO.Visible = False
    MEDIDA.Caption = "Unidade de Medida: lux"
    TIPOILUMINACAO.Visible = False
  End If
  'Ruido
  If (CurrentQuery.FieldByName("TIPO").AsInteger = 4) Then
    CALOR.Visible = False
    ILUMINANCIA.Visible = False
    TIPOILUMINACAO.Visible = False
    MEDIDA.Caption = "Unidade de Medida: dB (A)"
  End If
  'Geral/IluminaþÒo
  If (CurrentQuery.FieldByName("TIPO").AsInteger = 5) Then
    CALOR.Visible = False
    ILUMINANCIA.Visible = False
    RUIDO.Visible = False
    MEDIDA.Visible = False
    COMPONENTE.Visible = False
    TIPOILUMINACAO.Visible = True
    CurrentQuery.FieldByName("COMPONENTE").Value = "Iluminação"
  End If

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Sql As Object
  Dim TIPOR As Integer, ValorDigitado As Integer
  Dim Nivelaceitavel As String, nivelacao As String, nivelmonitoramento As String
  Dim Valor , Metade, Valor2

  If (CurrentQuery.FieldByName("TIPO").AsInteger = 4) Then
    If Not (CurrentQuery.FieldByName("MEDIDA").IsNull) Then
      Set Sql = NewQuery

      Sql.Add "SELECT  TIPORISCO, LIMITE, LIMITEACAO, LIMITEMONITORAMENTO FROM ST_RISCOS WHERE  HANDLE = 1 "
      Sql.Active = True
      TIPOR = Sql.FieldByName("TIPORISCO").Value
      If (TIPOR = 2) Or (TIPOR = 3) Then
        Valor = Sql.FieldByName("LIMITE").Value
        Nivelaceitavel = "NÝvel aceitßvel"
        nivelacao = "Nível de ação (>LT)"
        nivelmonitoramento = "Nível de ação (<LT)"

        ValorDigitado = CInt(CurrentQuery.FieldByName("MEDIDA").Value)

        If Not Sql.EOF And TIPOR = 2 Then
          If Valor > 0 Then
            Metade = Valor / 2
            If ValorDigitado < Metade Then
              CurrentQuery.FieldByName("SITUACAO").Value = Nivelaceitavel
              CurrentQuery.FieldByName("TIPOSITUACAO").Value = 3
            Else
              If ((ValorDigitado > Metade Or ValorDigitado = Metade) And ValorDigitado < Valor) Then
                CurrentQuery.FieldByName("SITUACAO").Value = nivelmonitoramento
                CurrentQuery.FieldByName("TIPOSITUACAO").Value = 2
              Else
                If (ValorDigitado > Valor Or ValorDigitado = Valor) Then
                  CurrentQuery.FieldByName("SITUACAO").Value = nivelacao
                  CurrentQuery.FieldByName("TIPOSITUACAO").Value = 1
                End If
              End If
            End If
          Else
            MsgBox "Risco sem valor do limite de tolerância"
            CanContinue = False
          End If
        Else
          Valor = Sql.FieldByName("LIMITEACAO").Value
          Valor2 = Sql.FieldByName("LIMITEMONITORAMENTO").Value
          If Not Sql.EOF And TIPOR = 3 Then
            If ((ValorDigitado > Valor2) Or (ValorDigitado = Valor2)) And (ValorDigitado < Valor) Then
              CurrentQuery.FieldByName("SITUACAO").Value = nivelmonitoramento
              CurrentQuery.FieldByName("TIPOSITUACAO").Value = 2
            Else
              If ((ValorDigitado > Valor) Or (ValorDigitado = Valor)) Then
                CurrentQuery.FieldByName("SITUACAO").Value = nivelacao
                CurrentQuery.FieldByName("TIPOSITUACAO").Value = 1
              Else
                CurrentQuery.FieldByName("SITUACAO").Value = " "
              End If
            End If
          End If
        End If
      Else
        MsgBox "Risco qualitativo"
        CanContinue = False
      End If
    End If
  End If
End Sub

Public Sub TIPOILUMINACAO_OnChange()
  If (CurrentQuery.FieldByName("TIPOILUMINACAO").AsInteger = 1) Then
    CurrentQuery.FieldByName("MEDIDA").Value = "Luz Natural"
  End If

  If (CurrentQuery.FieldByName("TIPOILUMINACAO").AsInteger = 2) Then
    CurrentQuery.FieldByName("MEDIDA").Value = "Luz Artificial Geral"
  End If

  If (CurrentQuery.FieldByName("TIPOILUMINACAO").AsInteger = 3) Then
    CurrentQuery.FieldByName("MEDIDA").Value = "Luz Artificial Suplementar"
  End If

  If (CurrentQuery.FieldByName("TIPOILUMINACAO").AsInteger = 4) Then
    CurrentQuery.FieldByName("MEDIDA").Value = "Luz Artificial Geral e Suplementar"
  End If

  If (CurrentQuery.FieldByName("TIPOILUMINACAO").AsInteger = 5) Then
    CurrentQuery.FieldByName("MEDIDA").Value = "Luz Natural e Artificial Geral"
  End If

  If (CurrentQuery.FieldByName("TIPOILUMINACAO").AsInteger = 6) Then
    CurrentQuery.FieldByName("MEDIDA").Value = "Luz Natural e Artificial Geral e Suplementar"
  End If
End Sub

