'HASH: 9C4C41F5CB4CB231C21AD34B144B794B
'Macro: EMPRESAS_COOPERATIVA
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not(CurrentQuery.FieldByName("PERCATOPRINCIPAL").AsFloat > 0) Then
    CanContinue = False
    bsShowMessage("Percentual do ato principal deve ser maior que 'Zero'", "E")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("PERCATOPRINCIPAL").AsFloat + _
      CurrentQuery.FieldByName("PERCATOAUXILIAR").AsFloat) <> 100 Then
    CanContinue = False
    bsShowMessage("A soma dos percentuais deve ser igual a 100%", "E")
  End If
End Sub

