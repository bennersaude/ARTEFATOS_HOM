'HASH: 2594A2C07709AA8E7E33B9EFFA116629

Public Sub ABERTURAPLANOFAIXA_OnChange()
  CurrentQuery.UpdateRecord

  If CurrentQuery.FieldByName("ABERTURAPLANOFAIXA").Value = "N" Then
    TIPOABERTURAPLANO.Visible = False
  Else
    TIPOABERTURAPLANO.Visible = True
  End If

  CurrentQuery.UpdateRecord
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("TIPOABRANGENCIA").Value = "1" Then
    REGIAO.Visible = False
  Else
    REGIAO.Visible = True
  End If

  If CurrentQuery.FieldByName("ABERTURAPLANOFAIXA").Value = "N" Then
    TIPOABERTURAPLANO.Visible = False
  Else
    TIPOABERTURAPLANO.Visible = True
  End If
End Sub

Public Sub TIPOABRANGENCIA_OnChange()
  CurrentQuery.UpdateRecord

  If CurrentQuery.FieldByName("TIPOABRANGENCIA").Value = "1" Then
    REGIAO.Visible = False
  Else
    REGIAO.Visible = True
  End If

  CurrentQuery.UpdateRecord
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	CanContinue = False

	If CurrentQuery.FieldByName("FAIXAVIDASMAXIMO").AsInteger < CurrentQuery.FieldByName("FAIXAVIDASMINIMO").AsInteger Then
		MsgBox("Máximo não pode ser menor que o mínimo!")
		FAIXAVIDASMAXIMO.SetFocus
		Exit Sub
	End If

	CanContinue = True
End Sub


