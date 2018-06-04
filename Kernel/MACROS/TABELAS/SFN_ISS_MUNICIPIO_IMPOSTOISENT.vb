'HASH: 74CA6C099CCC64F70CD09D965E2613F8
' ------------------------------ Macro Tabela SFN_ISS_MUNICIPIO_IMPOSTOISENTO -------------------------
'#Uses "*bsShowMessage"

' SMS 32908 - Kristian Fantin Pereira

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qCons As Object
  Set qCons = NewQuery

  'Verificação de existência de vigência aberta para a tesouraria
  qCons.Clear
  qCons.Add("SELECT COUNT(1) AS QTDE FROM SFN_ISS_MUNICIPIO_IMPOSTOISENT WHERE HANDLE <> :HANDLE AND DATAFINAL IS NULL")
  qCons.Add("   AND TESOURARIA = :TESOURARIA AND ISSMUNICIPIO = :ISSMUNICIPIO")
  qCons.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qCons.ParamByName("TESOURARIA").AsInteger = CurrentQuery.FieldByName("TESOURARIA").Value
  qCons.ParamByName("ISSMUNICIPIO").AsInteger = CurrentQuery.FieldByName("ISSMUNICIPIO").Value
  qCons.Active = True
  If qCons.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Cadastramento Impossível!" + Chr(13) + "Primeiro você precisar fechar a vigência atual em vigor", "E")

    qCons.Active = False
    Set qCons = Nothing
    CanContinue = False
    Exit Sub
  End If
  qCons.Active = False

  'Verificação de Data Válida
  If Not(CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    If CurrentQuery.FieldByName("DATAINICIAL").Value > CurrentQuery.FieldByName("DATAFINAL").Value Then
      bsShowMessage("A Data Final deve ser maior ou igual à Data Inicial", "E")

      CanContinue = False
      Exit Sub
    End If
  End If

  'Nao permitir cadastro de vigência retroativa
  qCons.Clear
  qCons.Add("SELECT MAX(DATAINICIAL) AS DATAMAIOR FROM SFN_ISS_MUNICIPIO_IMPOSTOISENT ")
  qCons.Add(" WHERE HANDLE <> :HANDLE AND TESOURARIA = :TESOURARIA AND ISSMUNICIPIO = :ISSMUNICIPIO")
  qCons.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qCons.ParamByName("TESOURARIA").AsInteger = CurrentQuery.FieldByName("TESOURARIA").Value
  qCons.ParamByName("ISSMUNICIPIO").AsInteger = CurrentQuery.FieldByName("ISSMUNICIPIO").Value
  qCons.Active = True
  If qCons.FieldByName("DATAMAIOR").AsDateTime > CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
    bsShowMessage("Cadastramento Impossível!" + Chr(13) + "Existe vigência com Data inicial maior a que você está tentando cadastrar!!!", "E")

    DATAINICIAL.SetFocus

    qCons.Active = False
    Set qCons = Nothing
    CanContinue = False
    Exit Sub
  End If
  qCons.Active = False

  'Verificação de não permitir cadastrar vigência dentro de intervalo existente
  qCons.Clear
  qCons.Add("SELECT COUNT(1) AS QTDE FROM SFN_ISS_MUNICIPIO_IMPOSTOISENT ")
  qCons.Add(" WHERE DATAINICIAL <= :DATA AND DATAFINAL >= :DATA AND HANDLE <> :HANDLE ")
  qCons.Add("   AND TESOURARIA = :TESOURARIA AND ISSMUNICIPIO = :ISSMUNICIPIO")
  qCons.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qCons.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").Value
  qCons.ParamByName("TESOURARIA").AsInteger = CurrentQuery.FieldByName("TESOURARIA").Value
  qCons.ParamByName("ISSMUNICIPIO").AsInteger = CurrentQuery.FieldByName("ISSMUNICIPIO").Value
  qCons.Active = True
  If qCons.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Não é possível Cadastrar uma vigência detro de um intervalo de Vigência existente!!!", "E")

    qCons.Active = False
    Set qCons = Nothing
    CanContinue = False
    Exit Sub
  End If
  qCons.Active = False

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    qCons.Clear
    qCons.Add("SELECT COUNT(1) AS QTDE FROM SFN_ISS_MUNICIPIO_IMPOSTOISENT ")
    qCons.Add(" WHERE DATAINICIAL BETWEEN :DATA1 AND :DATA2 AND HANDLE <> :HANDLE ")
    qCons.Add("   AND TESOURARIA = :TESOURARIA AND ISSMUNICIPIO = :ISSMUNICIPIO")
    qCons.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
    qCons.ParamByName("DATA1").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").Value
    qCons.ParamByName("DATA2").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").Value
    qCons.ParamByName("TESOURARIA").AsInteger = CurrentQuery.FieldByName("TESOURARIA").Value
    qCons.ParamByName("ISSMUNICIPIO").AsInteger = CurrentQuery.FieldByName("ISSMUNICIPIO").Value
    qCons.Active = True
    If qCons.FieldByName("QTDE").AsInteger > 0 Then
      bsShowMessage("Não é possível Fechar uma vigência que ultrapasse um intervalo de Vigência existente!!!", "E")

      qCons.Active = False
      Set qCons = Nothing
      CanContinue = False
      Exit Sub
    End If
    qCons.Active = False
  End If

  Set qCons = Nothing

End Sub


