'HASH: CB5B7B255E64798D990E3AA18CAAC994
'Macro TABELA SAM_ENTIDADE_TIPOCONTRIBUICAO (SMS 60044)

'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("VALORMINIMOCONTRIBUICAO").AsCurrency > CurrentQuery.FieldByName("VALORMAXIMOCONTRIBUICAO").AsCurrency Then
     bsShowMessage("Valor máximo de contribuição não pode ser menor que o valor mínimo de contribuição", "E")
     VALORMINIMOCONTRIBUICAO.SetFocus
     CanContinue = False
  End If

  If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
     If CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime > CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime Then
        bsShowMessage("Competência final não pode ser menor que a competência inicial", "E")
        COMPETENCIAINICIAL.SetFocus
        CanContinue = False
     End If
  End If

  '------------------------------------ Controle de Vigencia -------------------------------
  Dim qCons As Object
  Set qCons = NewQuery

  'Verificação de existência de vigência aberta
  qCons.Clear
  qCons.Add("SELECT COUNT(1) AS QTDE FROM SAM_ENTIDADE_TIPOCONTRIBUICAO WHERE HANDLE <> :HANDLE AND COMPETENCIAFINAL IS NULL AND ENTIDADE = :ENTIDADE AND TIPOCONTRIBUICAO = :TIPOCONTRIBUICAO")
  qCons.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qCons.ParamByName("ENTIDADE").AsInteger = CurrentQuery.FieldByName("ENTIDADE").Value
  qCons.ParamByName("TIPOCONTRIBUICAO").AsInteger = CurrentQuery.FieldByName("TIPOCONTRIBUICAO").Value
  qCons.Active = True
  If qCons.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Cadastramento Impossível!" + Chr(13) + "Primeiro você precisar fechar a competência atual em vigor", "E")

    qCons.Active = False
    Set qCons = Nothing
    CanContinue = False
    Exit Sub
  End If
  qCons.Active = False

  'Verificação de não permitir cadastrar competência dentro de intervalo existente
  qCons.Clear
  qCons.Add("SELECT COUNT(1) AS QTDE FROM SAM_ENTIDADE_TIPOCONTRIBUICAO")
  qCons.Add(" WHERE COMPETENCIAINICIAL <= :DATA AND COMPETENCIAFINAL >= :DATA")
  qCons.Add("   AND HANDLE <> :HANDLE AND ENTIDADE = :ENTIDADE AND TIPOCONTRIBUICAO = :TIPOCONTRIBUICAO")
  qCons.ParamByName("HANDLE").AsInteger           = CurrentQuery.FieldByName("HANDLE").Value
  qCons.ParamByName("DATA").AsDateTime            = CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime
  qCons.ParamByName("ENTIDADE").AsInteger         = CurrentQuery.FieldByName("ENTIDADE").Value
  qCons.ParamByName("TIPOCONTRIBUICAO").AsInteger = CurrentQuery.FieldByName("TIPOCONTRIBUICAO").Value
  qCons.Active = True
  If qCons.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Não é possível Cadastrar uma competência em um intervalo de competência existente!!!", "E")

    qCons.Active = False
    Set qCons = Nothing
    CanContinue = False
    Exit Sub
  End If
  qCons.Active = False

  'Nao permitir cadastro de competência retroativa
  qCons.Clear
  qCons.Add("SELECT MAX(COMPETENCIAINICIAL) AS DATAMAIOR")
  qCons.Add("  FROM SAM_ENTIDADE_TIPOCONTRIBUICAO       ")
  qCons.Add(" WHERE HANDLE <> :HANDLE                   ")
  qCons.Add("   AND ENTIDADE = :ENTIDADE AND TIPOCONTRIBUICAO = :TIPOCONTRIBUICAO ")

  qCons.ParamByName("HANDLE").AsInteger   = CurrentQuery.FieldByName("HANDLE").Value
  qCons.ParamByName("ENTIDADE").AsInteger = CurrentQuery.FieldByName("ENTIDADE").Value
  qCons.ParamByName("TIPOCONTRIBUICAO").AsInteger = CurrentQuery.FieldByName("TIPOCONTRIBUICAO").Value  
  qCons.Active = True
  If (qCons.FieldByName("DATAMAIOR").AsDateTime > CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime) And (CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull) Then
    bsShowMessage("Cadastramento Impossível!" + Chr(13) + "Existe competência com Data inicial maior a que você está tentando cadastrar!!!", "E")

    COMPETENCIAINICIAL.SetFocus

    qCons.Active = False
    Set qCons = Nothing
    CanContinue = False
    Exit Sub
  End If
  qCons.Active = False


  If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    qCons.Clear
    qCons.Add("SELECT COUNT(1) AS QTDE FROM SAM_ENTIDADE_TIPOCONTRIBUICAO")
    qCons.Add(" WHERE COMPETENCIAINICIAL BETWEEN :DATA1 AND :DATA2 AND HANDLE <> :HANDLE AND ENTIDADE = :ENTIDADE AND TIPOCONTRIBUICAO = :TIPOCONTRIBUICAO")
    qCons.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
    qCons.ParamByName("DATA1").AsDateTime = CurrentQuery.FieldByName("COMPETENCIAINICIAL").Value
    qCons.ParamByName("DATA2").AsDateTime = CurrentQuery.FieldByName("COMPETENCIAFINAL").Value
    qCons.ParamByName("ENTIDADE").AsInteger = CurrentQuery.FieldByName("ENTIDADE").Value
    qCons.ParamByName("TIPOCONTRIBUICAO").AsInteger = CurrentQuery.FieldByName("TIPOCONTRIBUICAO").Value
    qCons.Active = True
    If qCons.FieldByName("QTDE").AsInteger > 0 Then
      bsShowMessage("Não é possível Fechar uma competência que ultrapasse um intervalo de competência existente!!!", "E")

      qCons.Active = False
      Set qCons = Nothing
      CanContinue = False
      Exit Sub
    End If
    qCons.Active = False
  End If

  Set qCons = Nothing
  '------------------------------------- FIM CONTROLE DE VIGÊNCIA ----------------------------------------

End Sub
