'HASH: 89EC89ED3C0342B249B5763CB745CD77


Public Sub CONTRATOTPDEP_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  Dim Procura As Object
  Dim handlexx As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas = "SAM_TIPODEPENDENTE.DESCRICAO"

  vCriterio = "SAM_CONTRATO_TPDEP.CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO"))

  vCampos = "Tipo dependente"

  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_TPDEP|SAM_TIPODEPENDENTE[SAM_CONTRATO_TPDEP.TIPODEPENDENTE = SAM_TIPODEPENDENTE.HANDLE ]", vColunas, 1, vCampos, vCriterio, "Tipo de dependentes", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOTPDEP").Value = handlexx
  End If
  Set Procura = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		CONTRATOTPDEP.WebLocalWhere = "A.CONTRATO = " + CStr(RecordHandleOfTable("SAM_CONTRATO"))
	End If
End Sub
