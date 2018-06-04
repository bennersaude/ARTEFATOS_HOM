'HASH: 91551003CFD3B4693AD0B0BB53F5FE3A
'SAM_PACOTEAUXILIO

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vTotal As Double
  If (CurrentQuery.FieldByName("LIMITETIPO").AsString <> "S") And _
       (CurrentQuery.FieldByName("LIMITEVALOR").IsNull Or _
       (CurrentQuery.FieldByName("LIMITEVALOR").AsFloat <= 0)) Then
    CanContinue = False
    MsgBox("Deve ser informado um valor maior que Zero para o limite")
    Exit Sub
  End If
  'Balani SMS 49882 30/09/2005
  vTotal = CurrentQuery.FieldByName("PERCEMPRESA").AsFloat + CurrentQuery.FieldByName("PERCAUXILIO").AsFloat + CurrentQuery.FieldByName("PERCADIANTAMENTO").AsFloat
  'If (CurrentQuery.FieldByName("PERCEMPRESA").AsFloat + CurrentQuery.FieldByName("PERCAUXILIO").AsFloat + CurrentQuery.FieldByName("PERCADIANTAMENTO").AsFloat) <> 100 Then
  If (vTotal < 0.01) Or (vTotal > 100) Then
    CanContinue = False
    'MsgBox("A soma dos percentuais de Empresa, Auxílio e de Adiantamento deve ser igual a 100%")
    MsgBox("A soma dos percentuais de Empresa, Auxílio e Adiantamento devem estar entre 0.01% e 100%")
    Exit Sub
  End If

  'André - SMS 28326 - 15/10/2004
  '  Dim Linha As String
  '  Set Interface =CreateBennerObject("SAMGERAL.Vigencia")
  '  Linha =Interface.Vigencia(CurrentSystem,"SAM_PACOTEAUXILIO","DATAINICIAL","DATAFINAL",CurrentQuery.FieldByName("DATAINICIAL").AsDateTime,CurrentQuery.FieldByName("DATAFINAL").AsDateTime,"","")

  '  If Linha ="" Then
  '    CanContinue =True
  '  Else
  '    CanContinue =False
  '    MsgBox(Linha)
  '  End If
  'FIM SMS 28326

End Sub

