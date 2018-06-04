'HASH: A61CD8F31F65FC1384293DB4F04B7BAE

'#Uses "*bsShowMessage"
Option Explicit
Dim vdDataAdesao As Date
Dim viHBeneficiario As Long

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATAREATIVACAOESPECIAL").AsDateTime = ServerDate
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
  bs.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATAREATIVACAOESPECIAL").AsDateTime)
  bs.AddParameter(pdtAutomatic, vbReativarSomenteBenef)
  bs.AddParameter(pdtString, "E") 'Especial
  bsShowMessage(CStr(bs.Execute("Reativacao")), "I")
  Set bs = Nothing
End Sub
