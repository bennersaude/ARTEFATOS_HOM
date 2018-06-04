'HASH: F74460483BA80959592778589E9E776F
'#Uses "*bsShowMessage

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vDataAdesao As Integer
  Dim pSMensagem As String
  Dim gCampoDataAdesao As Date
  Dim vHPlano As Long
  Dim vHContrato As Long
  Dim vHModulo As Long

  vHPlano = CLng(SessionVar("HPLANO"))
  vHContrato = CLng(SessionVar("HCONTRATO"))
  vHModulo = CLng(SessionVar("HMODULO"))

  If OPCAODATA.PageIndex = 0 Then 'TAB = Data de adesão Do beneficiário
    vDataAdesao = 1
  Else
    vDataAdesao = 2 'TAB = outra data de adesão
    gCampoDataAdesao = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
  End If

  Set PropagaModulo = CreateBennerObject("Contrato.Propagar")
   	PropagaModulo.Exec(CurrentSystem, pSMensagem, vDataAdesao, gCampoDataAdesao, vHPlano, vHContrato, vHModulo)
    bsShowMessage(pSMensagem, "I")
  Set PropagaModulo = Nothing
End Sub
