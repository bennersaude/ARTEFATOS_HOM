'HASH: CE53EEA32620FA8C8FF4868647562BB7

' ------------------------------ Macro Tabela SFN_ISS_MUNICIPIO_REDUCAOCAM -------------------------
'#Uses "*bsShowMessage"

' SMS 33452 - Kristian Fantin Pereira

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qCons As Object
  Set qCons = NewQuery

  If CurrentQuery.FieldByName("EXERCICIOTRIBUTAVEL").AsInteger < 1 Then
    bsShowMessage("O Exercício Tributável deve ser maior que zero.", "E")
    EXERCICIOTRIBUTAVEL.SetFocus

    CanContinue = False
    Exit Sub
  End If

  'VERIFICACAO DA EXISTENCIA DE "EXERCICIOTRIBUTAVEL"
  qCons.Clear
  qCons.Add("SELECT COUNT(1) AS QTDE FROM SFN_ISS_MUNICIPIO_REDUCAOCAM WHERE HANDLE <> :HANDLE")
  qCons.Add("   AND ISSMUNICIPIO = :ISSMUNICIPIO AND EXERCICIOTRIBUTAVEL = :EXERCICIOTRIBUTAVEL")
  qCons.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qCons.ParamByName("ISSMUNICIPIO").AsInteger = CurrentQuery.FieldByName("ISSMUNICIPIO").Value
  qCons.ParamByName("EXERCICIOTRIBUTAVEL").AsInteger = CurrentQuery.FieldByName("EXERCICIOTRIBUTAVEL").Value
  qCons.Active = True
  If qCons.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Exercícío Tributável já existente. Favor corrigir!!!", "E")
    EXERCICIOTRIBUTAVEL.SetFocus

    qCons.Active = False
    Set qCons = Nothing
    CanContinue = False
    Exit Sub
  End If
  qCons.Active = False

  Set qCons = Nothing

End Sub


