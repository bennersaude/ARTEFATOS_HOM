'HASH: C35C0E8A279C1734153B7168CCC8FFFE

Public Sub TIPOABERTURA_OnChange()
  'If CurrentQuery.FieldByName("TIPOABERTURA").Value = "N" Then
  'If CurrentQuery.FieldByName("tipoabertura").Value = 1 Then
  '   ABERTURAPLANOFAIXA.Visible = False
  'Else
  '   ABERTURAPLANOFAIXA.Visible = True
  'End If


  'CurrentQuery.UpdateRecord

End Sub

Public Sub TIPOABRANGENCIA_OnChange()
  If CurrentQuery.FieldByName("TIPOABRANGENCIA").Value = "1" Then
    REGIAO.ReadOnly = False
  Else
    REGIAO.ReadOnly = True
  End If

  CurrentQuery.UpdateRecord
End Sub

Public Sub AberturaPlanoFaixa_OnChange()
  If CurrentQuery.FieldByName("ABERTURAPLANOFAIXA").AsBoolean = True Then
    TIPOABERTURA.Visible = False
    'TIPOABERTURA.ReadOnly = True
  Else
    TIPOABERTURA.Visible = True
    'TIPOABERTURA.ReadOnly = False
  End If

  CurrentQuery.UpdateRecord
End Sub

