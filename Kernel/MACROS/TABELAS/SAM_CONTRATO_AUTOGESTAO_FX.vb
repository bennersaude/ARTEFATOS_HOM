'HASH: 3CD1FA8650AADBD2D56234A938831DE2
'Macro da tabela: SAM_CONTRATO_AUTOGESTAO_FX
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Não permitir cadastrar duas faixas com mesma idade máxima e mesmo teto salarial.
  Dim qFaixasCruzadas As Object
  Set qFaixasCruzadas = NewQuery

  qFaixasCruzadas.Clear
  qFaixasCruzadas.Add("SELECT COUNT(1) QTD                            ")
  qFaixasCruzadas.Add("  FROM SAM_CONTRATO_AUTOGESTAO_FX              ")
  qFaixasCruzadas.Add(" WHERE CONTRATOAUTOGESTAO = :CONTRATOAUTOGESTAO")
  qFaixasCruzadas.Add("   AND TETOSALARIAL       = :TETOSALARIAL      ")
  qFaixasCruzadas.Add("   AND IDADEMAXIMA        = :IDADEMAXIMA       ")
  qFaixasCruzadas.Add("   AND HANDLE            <> :HANDLE            ")
  qFaixasCruzadas.ParamByName("CONTRATOAUTOGESTAO").AsInteger = CurrentQuery.FieldByName("CONTRATOAUTOGESTAO").AsInteger
  qFaixasCruzadas.ParamByName("TETOSALARIAL"      ).AsFloat   = CurrentQuery.FieldByName("TETOSALARIAL").AsFloat
  qFaixasCruzadas.ParamByName("IDADEMAXIMA"       ).AsInteger = CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger
  qFaixasCruzadas.ParamByName("HANDLE"            ).AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qFaixasCruzadas.Active = True

  If (qFaixasCruzadas.FieldByName("QTD").AsInteger > 0) Then
    Set qFaixasCruzadas = Nothing
    bsShowMessage("O teto salarial informado já está sendo utilizado em outra faixa etária com a mesma idade máxima.", "E")
    CanContinue = False
    Exit Sub
  End If
  Set qFaixasCruzadas = Nothing
End Sub
