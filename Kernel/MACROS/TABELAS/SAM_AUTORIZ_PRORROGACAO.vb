'HASH: B3F6B6BA48AEF0D06D771CABDAA7420E
'#uses "*bsShowMessage"'

Public Sub TABLE_AfterInsert()
  Dim qAutorizacao As Object
  Set qAutorizacao = NewQuery

  qAutorizacao.Clear
  qAutorizacao.Add("SELECT TISTIPOINTERNACAO")
  qAutorizacao.Add("FROM SAM_AUTORIZ")
  qAutorizacao.Add("WHERE HANDLE = :HANDLE")
  qAutorizacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
  qAutorizacao.Active = True

  If Not qAutorizacao.FieldByName("TISTIPOINTERNACAO").IsNull Then
    CurrentQuery.FieldByName("TIPOINTERNACAO").AsInteger = qAutorizacao.FieldByName("TISTIPOINTERNACAO").AsInteger
  End If
  Set qAutorizacao = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If WebMode Then
    If Not CurrentQuery.FieldByName("AUTORIZACAO").IsNull Then

      Dim qVerificaExistenciaProrrogacao As Object
      Set qVerificaExistenciaProrrogacao = NewQuery

      qVerificaExistenciaProrrogacao.Add("SELECT COUNT(1) QTD FROM SAM_AUTORIZ_PRORROGACAO WHERE AUTORIZACAO = :AUTORIZACAO AND PROTOCOLOTRANSACAO IS NULL AND HANDLE <> :HANDLE")
      qVerificaExistenciaProrrogacao.ParamByName("AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
      qVerificaExistenciaProrrogacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qVerificaExistenciaProrrogacao.Active = True

      If qVerificaExistenciaProrrogacao.FieldByName("QTD").AsInteger = 1 Then
        bsShowMessage("Já existe uma prorrogação em aberto para esta solicitação.", "E")
        Set qVerificaExistenciaProrrogacao = Nothing
        CanContinue = False
        Exit Sub
      End If

      Set qVerificaExistenciaProrrogacao = Nothing
    End If
  End If
End Sub

Public Sub TABLE_AfterScroll()

  WriteBDebugMessage("SAM_AUTORIZ_PRORROGACAO.TABLE_AfterScroll - Início")
  DATASOLICITACAO.ReadOnly = True
  Dim qAutorizacao As Object
  Set qAutorizacao = NewQuery

  qAutorizacao.Clear
  qAutorizacao.Add("SELECT BENEFICIARIO,")
  qAutorizacao.Add("       CONDICAOATENDIMENTO,")
  qAutorizacao.Add("       REGIMEINTERNACAO,")
  qAutorizacao.Add("       TIPOATENDIMENTO")
  qAutorizacao.Add("FROM SAM_AUTORIZ")
  qAutorizacao.Add("WHERE HANDLE = :HANDLE")
  qAutorizacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
  qAutorizacao.Active = True

  ACOMODACAOEVENTO.WebLocalWhere = " A.EVENTO IN (SELECT MODEV.EVENTO" + Chr(13) + _
                                   "              FROM SAM_MODULO_EVENTO MODEV" + Chr(13) + _
                                   "              JOIN SAM_CONTRATO_MOD CM ON CM.MODULO = MODEV.MODULO" + Chr(13) + _
                                   "              JOIN SAM_BENEFICIARIO_MOD BM ON CM.MODULO = CM.MODULO" + Chr(13) + _
                                   "              WHERE BM.BENEFICIARIO = " + qAutorizacao.FieldByName("BENEFICIARIO").AsString + Chr(13) + _
                                   "                AND BM.DATAADESAO <= @CAMPO(DATASOLICITACAO)" + Chr(13) + _
                                   "                AND (BM.DATACANCELAMENTO IS NULL OR" + Chr(13) + _
                                   "                     (BM.DATACANCELAMENTO IS NOT NULL AND BM.DATACANCELAMENTO >= @CAMPO(DATASOLICITACAO))))" + Chr(13) + _
                                   " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
                                   "                  FROM SAM_TGE TGE" + Chr(13) + _
                                   "                  JOIN SAM_CLASSEEVENTO_TIPOINTER CTI ON CTI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "                  WHERE TGE.HANDLE = A.EVENTO) OR " + Chr(13) + _
                                   "      EXISTS (SELECT 1" + Chr(13) + _
                                   "              FROM SAM_TGE TGE" + Chr(13) + _
                                   "              JOIN SAM_CLASSEEVENTO_TIPOINTER CTI ON CTI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "              WHERE TGE.HANDLE = A.EVENTO" + Chr(13) + _
                                   "                AND CTI.TIPOINTERNACAO = @CAMPO(TIPOINTERNACAO)))" + Chr(13) + _
                                   " AND A.ACOMODACAO IN (SELECT TIS.ACOMODACAO" + Chr(13) + _
                                   "                      FROM TIS_TIPOACOMODACAO TIS" + Chr(13) + _
                                   "                      WHERE TIS.VERSAOTISS = (SELECT MAX(HANDLE) VERSAOTISS FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'))"

  If Not qAutorizacao.FieldByName("CONDICAOATENDIMENTO").IsNull Then
    ACOMODACAOEVENTO.WebLocalWhere = ACOMODACAOEVENTO.WebLocalWhere + Chr(13) + " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
                                                                                "                  FROM SAM_TGE TGE" + Chr(13) + _
                                                                                "                  JOIN SAM_CLASSEEVENTO_CARATERATEND CCA ON CCA.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                                                                "                  WHERE TGE.HANDLE = A.EVENTO) OR " + Chr(13) + _
                                                                                "      EXISTS (SELECT 1" + Chr(13) + _
                                                                                "              FROM SAM_TGE TGE" + Chr(13) + _
                                                                                "              JOIN SAM_CLASSEEVENTO_CARATERATEND CCA ON CCA.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                                                                "              JOIN TIS_CARATERATENDIMENTO CAR ON CAR.HANDLE = CCA.CARATERATENDIMENTO" + Chr(13) + _
                                                                                "              WHERE TGE.HANDLE = A.EVENTO" + Chr(13) + _
                                                                                "                AND CAR.CONDICAOATENDIMENTO = " + qAutorizacao.FieldByName("CONDICAOATENDIMENTO").AsString + Chr(13) + _
                                                                                "                AND CAR.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')))"
  End If

  If Not qAutorizacao.FieldByName("REGIMEINTERNACAO").IsNull Then
    ACOMODACAOEVENTO.WebLocalWhere = ACOMODACAOEVENTO.WebLocalWhere + Chr(13) + " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
                                                                      "                  FROM SAM_TGE TGE" + Chr(13) + _
                                                                      "                  JOIN SAM_CLASSEEVENTO_REGIMEINTER CRI ON CRI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                                                      "                  WHERE TGE.HANDLE = A.EVENTO) OR" + Chr(13) + _
                                                                      "      EXISTS (SELECT 1" + Chr(13) + _
                                                                      "              FROM SAM_TGE TGE" + Chr(13) + _
                                                                      "              JOIN SAM_CLASSEEVENTO_REGIMEINTER CRI ON CRI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                                                      "              WHERE TGE.HANDLE = A.EVENTO" + Chr(13) + _
                                                                      "                AND CRI.REGIMEINTERNACAO = " + qAutorizacao.FieldByName("REGIMEINTERNACAO").AsString +"))"
  End If

  If Not qAutorizacao.FieldByName("TIPOATENDIMENTO").IsNull Then
    ACOMODACAOEVENTO.WebLocalWhere = ACOMODACAOEVENTO.WebLocalWhere + Chr(13) + " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
                                                                                "                  FROM SAM_TGE TGE" + Chr(13) + _
                                                                                "                  JOIN SAM_CLASSEEVENTO_TIPOATEND CTA ON CTA.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                                                                "                  WHERE TGE.HANDLE = A.EVENTO) OR " + Chr(13) + _
                                                                                "      EXISTS (SELECT 1" + Chr(13) + _
                                                                                "              FROM SAM_TGE TGE" + Chr(13) + _
                                                                                "              JOIN SAM_CLASSEEVENTO_TIPOATEND CTA ON CTA.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                                                                "              WHERE TGE.HANDLE = A.EVENTO" + Chr(13) + _
                                                                                "                AND CTA.TIPOATENDIMENTO = " + qAutorizacao.FieldByName("TIPOATENDIMENTO").AsString +"))"
  End If

  Set qAutorizacao = Nothing
  WriteBDebugMessage("SAM_AUTORIZ_PRORROGACAO.TABLE_AfterScroll - Fim")
End Sub
