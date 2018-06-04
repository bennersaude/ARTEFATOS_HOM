'HASH: 9ADC4C2B34E3B682BD9B6D9116F92648
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TABDOSEUNICA").AsInteger = 2 Then

    If CurrentQuery.FieldByName("DOSE2").AsInteger = 1 Then
      If CurrentQuery.FieldByName("TEMPORETORNODOSE2").AsInteger > 99999 Then
        CanContinue = False
        bsShowMessage("O tempo de retorno não pode ser superior à 99999: Dose 2.", "E")
        Exit Sub
      ElseIf CurrentQuery.FieldByName("TEMPORETORNODOSE2").AsInteger <= 0 Then
        CanContinue = False
        bsShowMessage("O tempo de retorno não pode ser inferior ou igual à zero: Dose 2.", "E")
        Exit Sub
      End If
    End If

    If CurrentQuery.FieldByName("DOSE3").AsInteger = 1 Then
      If CurrentQuery.FieldByName("TEMPORETORNODOSE3").AsInteger > 99999 Then
        CanContinue = False
        bsShowMessage("O tempo de retorno não pode ser superior à 99999: Dose 3.", "E")
        Exit Sub
      ElseIf CurrentQuery.FieldByName("TEMPORETORNODOSE3").AsInteger <= 0 Then
        CanContinue = False
        bsShowMessage("O tempo de retorno não pode ser inferior ou igual à zero: Dose 3.", "E")
        Exit Sub
      End If
    End If

    If CurrentQuery.FieldByName("REFORCO").AsInteger = 1 Then
      If CurrentQuery.FieldByName("TEMPORETORNOREFORCO").AsInteger > 99999 Then
        CanContinue = False
        bsShowMessage("O tempo de retorno não pode ser superior à 99999: Reforço.", "E")
        Exit Sub
      ElseIf CurrentQuery.FieldByName("TEMPORETORNOREFORCO").AsInteger <= 0 Then
        CanContinue = False
        bsShowMessage("O tempo de retorno não pode ser inferior ou igual à zero: Reforço.", "E")
        Exit Sub
      End If
    End If

  End If
End Sub
