'HASH: 8D8C447B6F9FB73DADB85179300C3AB9
'Macro: SAM_BENEFICIARIO_CARENCIA
'#Uses "*bsShowMessage"
Option Explicit


Public Sub CARENCIA_OnChange()
  If CurrentQuery.State = 3 Then ' Inclusão
    Dim qm As Object
    Set qm = NewQuery
    qm.Clear
    qm.Add("SELECT FC.QTDDIA")
    qm.Add("FROM SAM_BENEFICIARIO B, ")
    qm.Add("     SAM_FAMILIA_CARENCIA FC")
    qm.Add("WHERE B.HANDLE = :BENEFICIARIO")
    qm.Add(" AND FC.CARENCIA = :CONTRATOCARENCIA")
    qm.Add(" AND FC.FAMILIA = B.FAMILIA")
    qm.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    qm.ParamByName("CONTRATOCARENCIA").Value = CurrentQuery.FieldByName("CARENCIA").AsInteger
    qm.Active = True
    If Not qm.EOF Then
      If (Not qm.FieldByName("QTDDIA").IsNull) Then
        CurrentQuery.FieldByName("QTDDIA").Value = qm.FieldByName("QTDDIA").Value
      End If
    Else
      qm.Clear
      qm.Add("SELECT CC.QTDDIA")
      qm.Add("FROM SAM_CONTRATO_CARENCIA CC")
      qm.Add("WHERE CC.HANDLE = :CONTRATOCARENCIA")
      qm.ParamByName("CONTRATOCARENCIA").Value = CurrentQuery.FieldByName("CARENCIA").AsInteger
      qm.Active = True
      If Not qm.EOF Then
        If (Not qm.FieldByName("QTDDIA").IsNull) Then
          CurrentQuery.FieldByName("QTDDIA").Value = qm.FieldByName("QTDDIA").Value
        End If
      End If
    End If
    qm.Active = False
    Set qm = Nothing
  End If
End Sub


Public Sub TABLE_AfterInsert()
  Dim qBenef As Object
  Set qBenef = NewQuery



  qBenef.Clear
  qBenef.Add("SELECT SM.SEXO                                           ")
  qBenef.Add("  FROM SAM_MATRICULA    SM                               ")
  qBenef.Add("  JOIN SAM_BENEFICIARIO SB ON (SB.MATRICULA = SM.HANDLE) ")
  qBenef.Add(" WHERE SB.HANDLE = :BENEFICIARIO                         ")
  qBenef.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qBenef.Active = True


  If qBenef.FieldByName("SEXO").AsString = "F" Then
	    'SMS 77690 - ARTUR - Início - Inibe exibição de carencias inseridas no contrato para outros planos
	If WebMode Then
		CARENCIA.WebLocalWhere = "SEXO In ('F', 'A') And A.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @CAMPO(BENEFICIARIO)  And SBM.DATACANCELAMENTO Is Null GROUP BY scm.plano)"
	ElseIf VisibleMode Then
		CARENCIA.LocalWhere = "SEXO In ('F', 'A') And SAM_CONTRATO_CARENCIA.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @BENEFICIARIO  AND SBM.DATACANCELAMENTO IS NULL GROUP BY scm.plano)"
	End If
  Else
	If WebMode Then
	   	CARENCIA.WebLocalWhere = "SEXO In ('M', 'A') And A.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @CAMPO(BENEFICIARIO) AND SBM.DATACANCELAMENTO IS NULL GROUP BY scm.plano)"
	ElseIf VisibleMode Then
		CARENCIA.LocalWhere = "SEXO In ('M', 'A') And SAM_CONTRATO_CARENCIA.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @BENEFICIARIO AND SBM.DATACANCELAMENTO IS NULL GROUP BY scm.plano)"
	End If
	    'SMS 77690 - ARTUR - Fim - Inibe exibição de carencias inseridas no contrato para outros planos
  End If


  qBenef.Active = False
  Set qBenef = Nothing


  CurrentQuery.FieldByName("QTDDIA").Value = 0
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim qBenef As Object
  Set qBenef = NewQuery

  qBenef.Clear
  qBenef.Add("SELECT SM.SEXO                                           ")
  qBenef.Add("  FROM SAM_MATRICULA    SM                               ")
  qBenef.Add("  JOIN SAM_BENEFICIARIO SB ON (SB.MATRICULA = SM.HANDLE) ")
  qBenef.Add(" WHERE SB.HANDLE = :BENEFICIARIO                         ")
  qBenef.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qBenef.Active = True

  If qBenef.FieldByName("SEXO").AsString = "F" Then
	    'SMS 77690 - ARTUR - Início - Inibe exibição de carencias inseridas no contrato para outros planos
	If WebMode Then
		CARENCIA.WebLocalWhere = "SEXO In ('F', 'A') And A.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @CAMPO(BENEFICIARIO)  And SBM.DATACANCELAMENTO Is Null GROUP BY scm.plano)"
	ElseIf VisibleMode Then
		CARENCIA.LocalWhere = "SEXO In ('F', 'A') And SAM_CONTRATO_CARENCIA.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @BENEFICIARIO  AND SBM.DATACANCELAMENTO IS NULL GROUP BY scm.plano)"
	End If
  Else
	If WebMode Then
	   	CARENCIA.WebLocalWhere = "SEXO In ('M', 'A') And A.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @CAMPO(BENEFICIARIO) AND SBM.DATACANCELAMENTO IS NULL GROUP BY scm.plano)"
	ElseIf VisibleMode Then
		CARENCIA.LocalWhere = "SEXO In ('M', 'A') And SAM_CONTRATO_CARENCIA.PLANO In (Select scm.plano from sam_beneficiario sb Left Join sam_Beneficiario_mod sbm On sbm.beneficiario = sb.Handle Left Join sam_contrato_mod scm On scm.Handle = sbm.modulo where sb.Handle  = @BENEFICIARIO AND SBM.DATACANCELAMENTO IS NULL GROUP BY scm.plano)"
	End If
	    'SMS 77690 - ARTUR - Fim - Inibe exibição de carencias inseridas no contrato para outros planos
  End If

  qBenef.Active = False
  Set qBenef = Nothing
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT DATAADESAO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO")
  SQL.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("DATAADESAO").AsDateTime > CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
    bsShowMessage("Data inicial menor que a data de adesão do beneficiário.", "E")
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.State = 3 Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT * FROM SAM_BENEFICIARIO_CARENCIA C, SAM_BENEFICIARIO B ")
    SQL.Add(" WHERE C.BENEFICIARIO = B.HANDLE AND B.HANDLE = :BENEFICIARIO AND CARENCIA = :CARENCIA AND DATAINICIAL >= :DATAINICIAL ")
    SQL.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    SQL.ParamByName("CARENCIA").Value = CurrentQuery.FieldByName("CARENCIA").AsInteger
    SQL.ParamByName("DATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
    SQL.Active = True

    If Not SQL.EOF Then
      bsShowMessage("Data inicial inferior ou igual a uma data inicial já cadastrada.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If


  Set SQL = Nothing
End Sub

