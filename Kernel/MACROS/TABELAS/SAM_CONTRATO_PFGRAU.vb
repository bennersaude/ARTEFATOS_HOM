'HASH: F75E8BA8B4C6DF5FA3CB65B2C65E56EC
'Macro: SAM_CONTRATO_PFGRAU


'#Uses "*ProcuraGrau"
'#Uses "*bsShowMessage"
Dim gPlano As Long


Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  '  If Len(GRAU.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraGrau(GRAU.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  '  End If
End Sub


Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub


Public Sub PLANO_OnChange()
  'Anderson sms 21638
  gPlano = CurrentQuery.FieldByName("PLANO").AsInteger
End Sub


Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
  'Anderson sms 21638
  gPlano = CurrentQuery.FieldByName("PLANO").AsInteger
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
	ElseIf VisibleMode Then
		PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qContrato As Object
  Set qContrato = NewQuery

  qContrato.Clear
  qContrato.Add("SELECT DATAADESAO,     ")
  qContrato.Add("       DATACANCELAMENTO")
  qContrato.Add("  FROM SAM_CONTRATO    ")
  qContrato.Add(" WHERE HANDLE = :HANDLE")
  qContrato.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_CONTRATO")
  qContrato.Active = True

  'Não permitir a inclusão de registros com vigência iniciando antes da adesão do contrato ou terminando depois do cancelamento do mesmo (se houver).
  If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < qContrato.FieldByName("DATAADESAO").AsDateTime) Then
    bsShowMessage("A data inicial não pode ser anterior à adesão do contrato.", "E")
    CanContinue = False
    Set qContrato = Nothing
    Exit Sub
  Else
    If (Not qContrato.FieldByName("DATACANCELAMENTO").IsNull) Then
      If (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
        bsShowMessage("A data final não pode ficar em aberto, pois o contrato possui data de cancelamento.", "E")
        CanContinue = False
        Set qContrato = Nothing
        Exit Sub
      Else
        If (CurrentQuery.FieldByName("DATAFINAL").AsDateTime > qContrato.FieldByName("DATACANCELAMENTO").AsDateTime) Then
          bsShowMessage("A data final não pode ser posterior ao cancelamento do contrato.", "E")
          CanContinue = False
          Set qContrato = Nothing
          Exit Sub
        End If
      End If
    End If
  End If
  Set qContrato = Nothing

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  'Balani SMS 4595 17/08/2005
  If RecordHandleOfTable("SAM_CONTRATO_PFREGRA") > 0 Then
    Condicao = "AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString + " AND REGRA = " + Str(RecordHandleOfTable("SAM_CONTRATO_PFREGRA"))
  Else
    Condicao = "AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString + " AND REGRA IS NULL"'Anderson sms 21638
  End If
  'Final SMS 4595
  'Condicao ="AND GRAU = " +CurrentQuery.FieldByName("GRAU").AsString +" AND PLANO = " +CurrentQuery.FieldByName("PLANO").AsString 'Anderson sms 21638

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PFGRAU", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

End Sub

