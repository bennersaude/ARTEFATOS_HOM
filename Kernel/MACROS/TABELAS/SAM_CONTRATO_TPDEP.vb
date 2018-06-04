'HASH: CFDD30B8ACD61B01FFEB150FE1431C93
'Macro: SAM_CONTRATO_TPDEP
'#Uses "*bsShowMessage"
'Daniela Zardo -15/07/2002

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim qVerificaTPDEP As Object
  Set qVerificaTPDEP = NewQuery

  qVerificaTPDEP.Active = False
  qVerificaTPDEP.Clear
  qVerificaTPDEP.Add("SELECT Count(1) Encontrou ")
  qVerificaTPDEP.Add("  FROM SAM_CONTRATO_MODEVENTO_TPDEP ")
  qVerificaTPDEP.Add(" WHERE TIPODEPENDENTE = " + CurrentQuery.FieldByName("TIPODEPENDENTE").AsString )

  qVerificaTPDEP.Active = True

  If qVerificaTPDEP.FieldByName("Encontrou").AsInteger > 0 Then
    bsShowMessage("Não é possível a exclusão deste Tipo de Dependente pois existe relacionamento cadastrado no Tipo de Dependente do Evento do Módulo do Contrato!", "E")
    CanContinue = False
    Set qVerificaTPDEP = Nothing
    Exit Sub
  End If

  Set qVerificaTPDEP = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "AND TIPODEPENDENTE = " + CurrentQuery.FieldByName("TIPODEPENDENTE").AsString

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_TPDEP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  Set Interface = Nothing

  Dim SQL As Object
  Set SQL = NewQuery

  If CurrentQuery.State = 2 And _
                          Not(CurrentQuery.FieldByName("DATAFINAL").IsNull)Then

    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
      CanContinue = False
      bsShowMessage("Data final não pode ser inferior à data inicial", "E")
      Exit Sub
    End If

    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SAM_BENEFICIARIO")
    SQL.Add("WHERE TIPODEPENDENTE = :HCONTRATOTPDEP")
    SQL.Add("  AND DATAADESAO >= :DATAFINAL")

    SQL.ParamByName("HCONTRATOTPDEP").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.Active = True

    If Not SQL.EOF Then
      CanContinue = False
      Set SQL = Nothing
      bsShowMessage("Existem beneficiários cadastrados com adesão superior à data final. Alteração não permitida", "E")
    End If
  End If

  If CurrentQuery.State = 3 Or 2 Then
    If CurrentQuery.FieldByName("OBRIGATORIO").AsString = "S" Then
      Dim q1 As Object
      Set q1 = NewQuery

      q1.Add("SELECT HANDLE FROM SAM_CONTRATO WHERE HANDLE = :HCONTRATO AND DECOMPOSICAOFAMILIAR = 'S' ")
      q1.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")'("SAM_CONTRATO")
      q1.Active = True
      If q1.EOF Then
        bsShowMessage("Esse contrato não permite decomposição familiar", "E")
        CanContinue = False
      End If
    End If
  End If

  If Not CurrentQuery.FieldByName("QTDMAXIMA").IsNull Then
    Dim qVerificaQtdDepFam As Object
    Set qVerificaQtdDepFam = NewQuery

    qVerificaQtdDepFam.Clear
    qVerificaQtdDepFam.Add("SELECT FAMILIA, COUNT(1) QTDDEP")
    qVerificaQtdDepFam.Add("  FROM SAM_BENEFICIARIO")
    qVerificaQtdDepFam.Add("  WHERE CONTRATO = :CONTRATO AND")
    qVerificaQtdDepFam.Add("        TIPODEPENDENTE = :TIPODEPENDENTE  AND")
    qVerificaQtdDepFam.Add("        (  (  DATACANCELAMENTO IS NULL OR")
    qVerificaQtdDepFam.Add("              DATACANCELAMENTO >= :DATAATUAL) AND")
    qVerificaQtdDepFam.Add("           (  ATENDIMENTOATE IS NULL     OR")
    qVerificaQtdDepFam.Add("              ATENDIMENTOATE >= :DATAATUAL) )")
    qVerificaQtdDepFam.Add(" GROUP BY FAMILIA ")
    qVerificaQtdDepFam.Add(" HAVING COUNT(1) > :QTDMAXIMA")
    qVerificaQtdDepFam.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
    qVerificaQtdDepFam.ParamByName("QTDMAXIMA").Value = CurrentQuery.FieldByName("QTDMAXIMA").AsInteger
    qVerificaQtdDepFam.ParamByName("TIPODEPENDENTE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    qVerificaQtdDepFam.ParamByName("DATAATUAL").AsDateTime = ServerDate
    qVerificaQtdDepFam.Active = True

    If qVerificaQtdDepFam.FieldByName("FAMILIA").AsInteger > 0 Then
      CanContinue = False
      bsShowMessage("Quantidade Máxima Inválida!", "E")
      Set qVerificaQtdDepFam = Nothing
      Exit Sub
    End If

    Set qVerificaQtdDepFam = Nothing
  End If

End Sub

