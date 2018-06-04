'HASH: DED9C84AAB9E51351079DC393AD3FFED
'#Uses "*bsShowMessage"

Public Function CHECARCARENCIA()
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object

  CHECARCARENCIA = True

  Condicao = " AND CARENCIA = " + CStr(CurrentQuery.FieldByName("carencia").AsInteger) + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString 'Anderson sms 21638(PLANO)

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_CARENCIA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    CHECARCARENCIA = False
  Else
    CHECARCARENCIA = True
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

End Function


Public Sub BOTAODUPLICARCARENCIA_OnClick()
  Dim Interface As Object

  'Daniela Zardo -18/07/2002
  Set Interface = CreateBennerObject("CONTRATO.DuplicarCarencia")
  Interface.DuplicarCarencia(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("CONTRATO").AsInteger, CurrentQuery.FieldByName("CARENCIA").AsInteger, CurrentQuery.FieldByName("QTDDIA").AsInteger)
  Set Interface = Nothing

End Sub

Public Sub CARENCIA_OnChange()
  If CurrentQuery.State = 3 Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Add("SELECT A.*                   ")
    SQL.Add("  FROM SAM_PLANO_CARENCIA  A,")
    SQL.Add("       SAM_CONTRATO        B ")
    SQL.Add("WHERE B.HANDLE = :HCONTRATO  ")
    SQL.Add("  AND A.PLANO = B.PLANO      ")'Anderson sms 21638
    SQL.Add("  AND A.CARENCIA = :HCARENCIA")
    SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
    SQL.ParamByName("HCARENCIA").Value = CurrentQuery.FieldByName("CARENCIA").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      CurrentQuery.FieldByName("TABREGRAREDE").Value = SQL.FieldByName("TABREGRAREDE").AsInteger
      CurrentQuery.FieldByName("QTDDIA").Value = SQL.FieldByName("QTDDIA").AsInteger
      CurrentQuery.FieldByName("TABREGRAREDEPROPRIA").Value = SQL.FieldByName("TABREGRAREDEPROPRIA").AsInteger
      CurrentQuery.FieldByName("QTDDIASREDEPROPRIA").Value = SQL.FieldByName("QTDDIASREDEPROPRIA").AsInteger
    End If
    Set SQL = Nothing
  End If
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT PLANO FROM SAM_CONTRATO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_CONTRATO")
  SQL.Active = True

  If WebMode Then
  	  CARENCIA.WebLocalWhere = "HANDLE IN (SELECT CARENCIA FROM SAM_PLANO_CARENCIA WHERE PLANO = @CAMPO(PLANO))"
  	  PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
  ElseIf VisibleMode Then
	  CARENCIA.LocalWhere = "HANDLE IN (SELECT CARENCIA FROM SAM_PLANO_CARENCIA WHERE PLANO = @PLANO)"
	  PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
  End If

  SQL.Active = False
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT PLANO FROM SAM_CONTRATO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_CONTRATO")
  SQL.Active = True

  If WebMode Then
  	  CARENCIA.WebLocalWhere = "HANDLE IN (SELECT CARENCIA FROM SAM_PLANO_CARENCIA WHERE PLANO = @CAMPO(PLANO))"
  	  PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
  ElseIf VisibleMode Then
	  CARENCIA.LocalWhere = "HANDLE IN (SELECT CARENCIA FROM SAM_PLANO_CARENCIA WHERE PLANO = @PLANO)"
	  PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
  End If

  SQL.Active = False
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  CanContinue = True
  If CurrentQuery.State = 3 Then
    SQL.Add("SELECT DATAINICIAL ")
    SQL.Add("FROM SAM_CONTRATO_CARENCIA ")
    SQL.Add("WHERE CONTRATO   = :HCONTRATO")
    SQL.Add("  AND PLANO = :PLANO")'Anderson sms 21638
    SQL.Add("  AND CARENCIA = :HCARENCIA")
    SQL.Add("  AND DATAFINAL IS NULL")
    SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
    SQL.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger 'Anderson sms 21638
    SQL.ParamByName("HCARENCIA").Value = CurrentQuery.FieldByName("CARENCIA").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      bsShowMessage("Carência tem vigência em aberto", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If
  Set SQL = Nothing
  If CHECARCARENCIA Then
    CanContinue = False
    Exit Sub
  End If
  If CurrentQuery.FieldByName("TABREGRAREDE").AsInteger = 2 Then
    CurrentQuery.FieldByName("QTDDIA").Clear
  End If
  If CurrentQuery.FieldByName("TABREGRAREDEPROPRIA").AsInteger = 2 Then
    CurrentQuery.FieldByName("QTDDIASREDEPROPRIA").Clear
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAODUPLICARCARENCIA" Then
		BOTAODUPLICARCARENCIA_OnClick
	End If
End Sub
