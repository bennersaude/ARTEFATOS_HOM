'HASH: 72CCC2FEB504A37E03F8627FC688100E
 
'#uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TABTIPOFORMATO").AsInteger = 1 Then
    If CurrentQuery.FieldByName("COLUNADOCUMENTO").AsInteger > CurrentQuery.FieldByName("TOTALCOLUNAS").AsInteger Then
      CanContinue = False
      bsShowMessage("Coluna do ""Documento"" não pode ser superior ao total de colunas!", "E")
    End If

    If CurrentQuery.FieldByName("COLUNADATALANCAMENTO").AsInteger > CurrentQuery.FieldByName("TOTALCOLUNAS").AsInteger Then
      CanContinue = False
      bsShowMessage("Coluna da ""Data do lançamento"" não pode ser superior ao total de colunas!", "E")
    End If

    If CurrentQuery.FieldByName("COLUNAHISTORICO").AsInteger > CurrentQuery.FieldByName("TOTALCOLUNAS").AsInteger Then
      CanContinue = False
      bsShowMessage("Coluna do ""Histórico"" não pode ser superior ao total de colunas!", "E")
    End If

    If CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").AsInteger > CurrentQuery.FieldByName("TOTALCOLUNAS").AsInteger Then
      CanContinue = False
      bsShowMessage("Coluna do ""Histórico complementar"" não pode ser superior ao total de colunas!", "E")
    End If

    If CurrentQuery.FieldByName("COLUNANATUREZA").AsInteger > CurrentQuery.FieldByName("TOTALCOLUNAS").AsInteger Then
      CanContinue = False
      bsShowMessage("Coluna da ""Natureza"" não pode ser superior ao total de colunas!", "E")
    End If

    If CurrentQuery.FieldByName("COLUNAVALOR").AsInteger > CurrentQuery.FieldByName("TOTALCOLUNAS").AsInteger Then
      CanContinue = False
      bsShowMessage("Coluna do ""Valor"" não pode ser superior ao total de colunas!", "E")
    End If

    If CurrentQuery.FieldByName("IDENTIFICADORNATUREZA").AsString = "3" And _
       Not CurrentQuery.FieldByName("COLUNANATUREZA").IsNull Then
      CanContinue = False
      bsShowMessage("Quando o campo ""Indicador da natureza do lançamento"" estiver configurado com a opção ""Determinado pelo valor"" não se deve preencher a coluna da natureza!", "E")
    End If

    If CurrentQuery.FieldByName("IDENTIFICADORNATUREZA").AsString <> "3" And _
       CurrentQuery.FieldByName("COLUNANATUREZA").IsNull Then
      CanContinue = False
      bsShowMessage("Preenchimento do campo ""Natureza"" é obrigatório!", "E")
    End If

    If (CurrentQuery.FieldByName("COLUNADOCUMENTO").AsInteger = CurrentQuery.FieldByName("COLUNADATALANCAMENTO").AsInteger Or _
        CurrentQuery.FieldByName("COLUNADOCUMENTO").AsInteger = CurrentQuery.FieldByName("COLUNAHISTORICO").AsInteger Or _
        CurrentQuery.FieldByName("COLUNADOCUMENTO").AsInteger = CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").AsInteger Or _
        CurrentQuery.FieldByName("COLUNADOCUMENTO").AsInteger = CurrentQuery.FieldByName("COLUNANATUREZA").AsInteger Or _
        CurrentQuery.FieldByName("COLUNADOCUMENTO").AsInteger = CurrentQuery.FieldByName("COLUNAVALOR").AsInteger) Or _
       (CurrentQuery.FieldByName("COLUNADATALANCAMENTO").AsInteger = CurrentQuery.FieldByName("COLUNAHISTORICO").AsInteger Or _
        CurrentQuery.FieldByName("COLUNADATALANCAMENTO").AsInteger = CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").AsInteger Or _
        CurrentQuery.FieldByName("COLUNADATALANCAMENTO").AsInteger = CurrentQuery.FieldByName("COLUNANATUREZA").AsInteger Or _
        CurrentQuery.FieldByName("COLUNADATALANCAMENTO").AsInteger = CurrentQuery.FieldByName("COLUNAVALOR").AsInteger) Or _
       (CurrentQuery.FieldByName("COLUNAHISTORICO").AsInteger = CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").AsInteger Or _
        CurrentQuery.FieldByName("COLUNAHISTORICO").AsInteger = CurrentQuery.FieldByName("COLUNANATUREZA").AsInteger Or _
        CurrentQuery.FieldByName("COLUNAHISTORICO").AsInteger = CurrentQuery.FieldByName("COLUNAVALOR").AsInteger) Or _
       (Not(CurrentQuery.FieldByName("COLUNANATUREZA").IsNull) And _
        Not(CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").IsNull) And _
        (CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").AsInteger = CurrentQuery.FieldByName("COLUNANATUREZA").AsInteger)) Or _
       (Not(CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").IsNull) And _
        (CurrentQuery.FieldByName("COLUNAHISTORICOCOMPLEMENTAR").AsInteger = CurrentQuery.FieldByName("COLUNAVALOR").AsInteger)) Then
      CanContinue = False
      bsShowMessage("Não pode existir mais de um campo para uma mesma coluna!", "E")
    End If
  End If
End Sub
