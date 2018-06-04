'HASH: 76D9BCC3248658253CDCE7B3633CB648
'#Uses "*bsShowMessage"
Dim viHBeneficiario As Long
Dim vdDataAdesao As Date

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
End Sub

Public Sub TABLE_AfterScroll()

  If SessionVar("HANDLECA019") <> "" Then
	viHBeneficiario = CLng(SessionVar("HANDLECA019"))
  Else
  	viHBeneficiario = CLng(SessionVar("HBENEFICIARIO"))
  End If

  Dim qBenef As Object
  Set qBenef = NewQuery

  qBenef.Add("SELECT DATAADESAO,            ")
  qBenef.Add("       EHTITULAR              ")
  qBenef.Add("  FROM SAM_BENEFICIARIO       ")
  qBenef.Add(" WHERE HANDLE = :HBENEFICIARIO")
  qBenef.ParamByName("HBENEFICIARIO").AsInteger = viHBeneficiario
  qBenef.Active = True

  vdDataAdesao = qBenef.FieldByName("DATAADESAO").AsDateTime

  ' SMS: 339622 - verifica se é titular
  If (qBenef.FieldByName("EHTITULAR").Value = "S") Then
    REATIVARSOMENTEBENEF.Visible = True
  Else
    REATIVARSOMENTEBENEF.Visible = False
  End If

  Set qBenef = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim bs As CSBusinessComponent
  Dim vbReativarSomenteBenef As Boolean

  ' SMS: 339622 - verificar se o flag 'REATIVARSOMENTEBENEF'
  If (CurrentQuery.FieldByName("REATIVARSOMENTEBENEF").Value = "S") Then
    vbReativarSomenteBenef = True

  Else
    vbReativarSomenteBenef = False

  End If

  Set bs = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.Beneficiarios.Cancelamento, Benner.Saude.Beneficiarios.Business") ' formato: [namespace.classe], [assembly]

  bs.ClearParameters
  bs.AddParameter(pdtInteger, viHBeneficiario)
  bs.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime)
  bs.AddParameter(pdtAutomatic, vbReativarSomenteBenef)
  bs.AddParameter(pdtString, "N") 'Normal
  bsShowMessage(CStr(bs.Execute("Reativacao")), "I")
  Set bs = Nothing

End Sub
