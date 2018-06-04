'HASH: 9FCA452B512188454DF1E7F76B6EBC20
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebVisionCode = "V_SAM_CONTRATO_PFEVENTO_IDADE" Then
			CONTRATOPFEVENTO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("IDADEMAXIMA").AsFloat < CurrentQuery.FieldByName("IDADEMINIMA").AsFloat Then
    bsShowMessage("Idade máxima deve ser maior que a idade mínima.", "E")
    CanContinue = False
    Exit Sub
  End If
  Dim qFaixa As Object
  Set qFaixa = NewQuery
  qFaixa.Active = False
  qFaixa.Clear
  qFaixa.Add("SELECT COUNT(1) QTDE")
  qFaixa.Add("  FROM SAM_CONTRATO_PFEVENTO_IDADE")
  qFaixa.Add(" WHERE CONTRATOPFEVENTO = :CONTRATOPFEVENTO")
  qFaixa.Add("   AND (")
  qFaixa.Add("        (:IDADEMINIMA BETWEEN IDADEMINIMA AND IDADEMAXIMA)")
  qFaixa.Add("        OR (:IDADEMAXIMA BETWEEN IDADEMINIMA AND IDADEMAXIMA)")
  qFaixa.Add("        OR (:IDADEMINIMA <= IDADEMINIMA AND :IDADEMAXIMA >= IDADEMAXIMA)")
  qFaixa.Add("       )")
  qFaixa.Add("        AND HANDLE <> :HANDLE")
  qFaixa.ParamByName("CONTRATOPFEVENTO").AsInteger = CurrentQuery.FieldByName("CONTRATOPFEVENTO").AsInteger
  qFaixa.ParamByName("IDADEMINIMA").AsFloat = CurrentQuery.FieldByName("IDADEMINIMA").AsFloat
  qFaixa.ParamByName("IDADEMAXIMA").AsFloat = CurrentQuery.FieldByName("IDADEMAXIMA").AsFloat
  qFaixa.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qFaixa.Active = True
  If qFaixa.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Não é permitido incluir duas faixas cruzadas.", "E")
    CanContinue = False
  End If
  Set qFaixa = Nothing
End Sub

