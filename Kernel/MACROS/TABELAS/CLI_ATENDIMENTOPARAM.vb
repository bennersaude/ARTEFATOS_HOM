'HASH: F8F5EB95392439C1B2A68DB33FD831EB
'MACRO: CLI_ATENDIMENTOPARAM

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("ESPECIFICACAOATENDIMENTO").AsString <> "P" Then
    If CurrentQuery.FieldByName("AUDIOMETRIA").AsString = "S" Then
      bsShowMessage("A opção 'Audiometria' só pode ser selecionada para PCMSO!", "E")
      CanContinue = False
    End If

    If CurrentQuery.FieldByName("AVALIACAOVISUAL").AsString = "S" Then
      bsShowMessage("A opção 'Avaliação visual' só pode ser selecionada para PCMSO!", "E")
      CanContinue = False
    End If

    If CurrentQuery.FieldByName("ESPIROMETRIA").AsString = "S" Then
      bsShowMessage("A opção 'Espirometria' só pode ser selecionada para PCMSO!", "E")
      CanContinue = False
    End If

    If CurrentQuery.FieldByName("OFTALMOLOGIA").AsString = "S" Then
      bsShowMessage("A opção 'Oftalmologia' só pode ser selecionada para PCMSO!", "E")
      CanContinue = False
    End If

    If CurrentQuery.FieldByName("EXAMECLINICOGERAL").AsString = "S" Then
      bsShowMessage("A opção 'Exame clínico geral' só pode ser selecionada para PCMSO!", "E")
      CanContinue = False
    End If

    If CurrentQuery.FieldByName("FORMULARIOS").AsString = "S" Then
      bsShowMessage("A opção 'Formulários' só pode ser selecionada para PCMSO!", "E")
      CanContinue = False
    End If

  End If
End Sub
