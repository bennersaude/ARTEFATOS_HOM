'HASH: 52C400764111C386BA9E6D203AF341A5
'#Uses "*bsShowMessage"
'SAM_CONTRATO_TPDEP_LIMITACAO

Public Sub CONTRATOLIMITACAO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long
  Dim SQL As Object


  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_LIMITACAO|SAM_LIMITACAO[SAM_CONTRATO_LIMITACAO.LIMITACAO = SAM_LIMITACAO.HANDLE]", "DESCRICAO", 1, "Descrição", "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO")), "Procura por Limitação", True, "")
  If handlexx <= 0 Then
    Exit Sub
  Else
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOLIMITACAO").Value = handlexx

    Set SQL = NewQuery
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT CONTRATOLIMITACAO                     ")
    SQL.Add("  FROM SAM_CONTRATO_TPDEP_LIMITACAO          ")
    SQL.Add(" WHERE CONTRATOLIMITACAO = :CONTRATOLIMITACAO")
    SQL.Add("   AND HANDLE <> :HANDLE                     ")
    SQL.ParamByName("CONTRATOLIMITACAO").Value = handlexx
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If Not SQL.FieldByName("CONTRATOLIMITACAO").IsNull Then
      bsShowMessage("Esta limitação já está cadastrado.", "I")
      Exit Sub
    End If

    'SMS 61198 - Matheus - Início
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT SL.PERIODICIDADE          ")
    SQL.Add("  FROM SAM_LIMITACAO SL,         ")
    SQL.Add("       SAM_CONTRATO_LIMITACAO SCL")
    SQL.Add(" WHERE SCL.HANDLE = :HANDLE      ")
    SQL.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATOLIMITACAO").AsInteger
    SQL.Active = True

    If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then
      CurrentQuery.FieldByName("PERIODO").AsInteger = 1
      PERIODO.Visible = False
    Else
      CurrentQuery.FieldByName("PERIODO").Clear
      PERIODO.Visible = True
    End If

    Set SQL = Nothing
    'SMS 61198 - Matheus - Fim

 End If
  Set Procura = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  'SMS 61198 - Matheus - Início
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SL.PERIODICIDADE          ")
  SQL.Add("  FROM SAM_LIMITACAO SL,         ")
  SQL.Add("       SAM_CONTRATO_LIMITACAO SCL")
  SQL.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  SQL.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATOLIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODO.Visible = False
  Else
    PERIODO.Visible = True
  End If

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'SMS 61198 - Matheus - Início
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SL.PERIODICIDADE          ")
  SQL.Add("  FROM SAM_LIMITACAO SL,         ")
  SQL.Add("       SAM_CONTRATO_LIMITACAO SCL")
  SQL.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  SQL.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATOLIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then  CurrentQuery.FieldByName("PERIODO").AsInteger = 1

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub
