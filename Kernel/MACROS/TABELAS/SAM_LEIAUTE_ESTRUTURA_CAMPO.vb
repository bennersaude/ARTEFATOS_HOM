'HASH: 5C7580B739EB00E5CE42CFBD8981DA11
'Tabela: SAM_LEIAUTE_ESTRUTURA_CAMPO
'Criada em 22/01/2002
'SMS 5694 - Milton

'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim q1 As Object
  Set q1 = NewQuery

  If (CurrentQuery.FieldByName("TAMANHO").AsInteger < 1) And (Not CurrentQuery.FieldByName("TAMANHO").IsNull) Then
  'Tamanho zero ou negativo
  bsShowMessage("Tamanho deve ser maior que zero", "E")
  CanContinue = False
  Exit Sub
  TAMANHO.SetFocus
  'Else
  '  If CurrentQuery.FieldByName("POSICAOINICIAL").IsNull Then
  '    If CurrentQuery.FieldByName("POSICAOFINAL").IsNull Then
  '      'só tem tamanho
  '      q1.Clear
  '      q1.Add("SELECT CASE WHEN MAX(POSICAOFINAL) IS NULL THEN 0 ELSE MAX(POSICAOFINAL) END MAIORPOSICAOFINAL FROM SAM_LEIAUTE_ESTRUTURA_CAMPO")
  '      q1.Add("WHERE LEIAUTEGRUPO=:ESTRUTURA AND HANDLE<>:HANDLE")
  '      q1.ParamByName("ESTRUTURA").Value = CurrentQuery.FieldByName("LEIAUTEGRUPO").Value
  '      q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  '      q1.Active = True
  '      'calcula as posições, conforme o tamanho
  '      CurrentQuery.FieldByName("POSICAOINICIAL").Value = (q1.FieldByName("MAIORPOSICAOFINAL").Value) + 1
  '      CurrentQuery.FieldByName("POSICAOFINAL").Value = (CurrentQuery.FieldByName("POSICAOINICIAL").Value) + (CurrentQuery.FieldByName("TAMANHO").Value) -1
  '    End If
  '  End If
  End If

  If CurrentQuery.FieldByName("TAMANHO").IsNull Then
    If CurrentQuery.FieldByName("POSICAOINICIAL").IsNull Then
      If CurrentQuery.FieldByName("POSICAOFINAL").IsNull Then
        'Tudo nulo
        bsShowMessage("Informe Tamanho ou Posições Inicial e Final", "E")
        CanContinue = False
        TAMANHO.SetFocus
      Else
        'Só tem posição final
        bsShowMessage("Informe Tamanho ou Posição Inicial", "E")
        CanContinue = False
        TAMANHO.SetFocus
      End If
    Else
      If CurrentQuery.FieldByName("POSICAOFINAL").IsNull Then
        'só tem posição inicial
        bsShowMessage("Informe Tamanho ou Posição Final", "E")
        CanContinue = False
        TAMANHO.SetFocus
      Else
        'tem posição inicial e final
        CurrentQuery.FieldByName("TAMANHO").Value = (CurrentQuery.FieldByName("POSICAOFINAL").Value) - (CurrentQuery.FieldByName("POSICAOINICIAL").Value) + 1
        CanContinue = True
      End If
    End If
  Else 'tamanho não é nulo
    If CurrentQuery.FieldByName("POSICAOINICIAL").IsNull Then
      If CurrentQuery.FieldByName("POSICAOFINAL").IsNull Then
        'só tem tamanho
        q1.Clear
        q1.Add("SELECT CASE WHEN MAX(POSICAOFINAL) IS NULL THEN 0 ELSE MAX(POSICAOFINAL) END MAIORPOSICAOFINAL FROM SAM_LEIAUTE_ESTRUTURA_CAMPO")
        q1.Add("WHERE LEIAUTEGRUPO=:ESTRUTURA AND HANDLE<>:HANDLE")
        q1.ParamByName("ESTRUTURA").Value = CurrentQuery.FieldByName("LEIAUTEGRUPO").Value
        q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
        q1.Active = True
        'calcula as posições, conforme o tamanho
        CurrentQuery.FieldByName("POSICAOINICIAL").Value = (q1.FieldByName("MAIORPOSICAOFINAL").Value) + 1
        CurrentQuery.FieldByName("POSICAOFINAL").Value = (CurrentQuery.FieldByName("POSICAOINICIAL").Value) + (CurrentQuery.FieldByName("TAMANHO").Value) -1
        CanContinue = True
      Else
        'tem posição final e tamanho
        bsShowMessage("Informe a Posição Inicial ou apague a Posição Final antes de salvar", "E")
        CanContinue = False
      End If
    Else
      If CurrentQuery.FieldByName("POSICAOFINAL").IsNull Then
        'tem tamanho e Pos. Inicial
        CurrentQuery.FieldByName("POSICAOFINAL").Value = (CurrentQuery.FieldByName("POSICAOINICIAL").Value) + (CurrentQuery.FieldByName("TAMANHO").Value) -1
        CanContinue = True
      Else
        'tem todos
        If CurrentQuery.FieldByName("TAMANHO").Value = (CurrentQuery.FieldByName("POSICAOFINAL").Value) - (CurrentQuery.FieldByName("POSICAOINICIAL").Value) + 1 Then
          CanContinue = True
        Else
          bsShowMessage("Posições inicial e final não compatíveis com o tamanho do campo.", "E")
          CanContinue = False
          TAMANHO.SetFocus
        End If
      End If
    End If
  End If


End Sub

