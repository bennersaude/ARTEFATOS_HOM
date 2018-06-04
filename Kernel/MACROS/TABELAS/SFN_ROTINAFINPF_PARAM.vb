'HASH: AFC7183448CBCFBD652169B4B7F693AD
'#Uses "*bsShowMessage"
'#Uses "*ProcuraBeneficiarioAtivo"

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False

  vHandle = ProcuraBeneficiarioAtivo(False, ServerDate, BENEFICIARIO.Text)
  CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle


End Sub

Function VerificaSeProcessada As Boolean

  Dim SQLRotFin As Object
  Dim HandleRotinaFinPF As Long
  Set SQLRotFin = NewQuery

  HandleRotinaFinPF = RecordHandleOfTable("SFN_ROTINAFINPF")

  SQLRotFin.Add("SELECT A.ROTINAFIN, B.SITUACAO")
  SQLRotFin.Add("  FROM SFN_ROTINAFINPF A ,")
  SQLRotFin.Add("       SFN_ROTINAFIN B")
  SQLRotFin.Add(" WHERE A.HANDLE = :ROTINAFINPF")
  SQLRotFin.Add("   AND B.HANDLE = A.ROTINAFIN")
  SQLRotFin.ParamByName("ROTINAFINPF").Value = HandleRotinaFinPF
  SQLRotFin.Active = True

  VerificaSeProcessada = SQLRotFin.FieldByName("SITUACAO").Value = "P"

  SQLRotFin.Active = False

  Set SQLRotFin = Nothing

End Function

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If VerificaSeProcessada Then
	  	bsShowMessage("Exclusão não permitida. Rotina já processada.","E")
	  	CanContinue = False
	  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If VerificaSeProcessada Then
		bsShowMessage("Alteração não permitida. Rotina já processada.","E")
		CanContinue = False
	End If
End Sub


Public Sub CONTRATOFINAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND EMPRESA = " + Str(CurrentCompany)
  vCriterio = vCriterio + "AND TABTIPOCONTRATO = 1"
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, 2, vCampos, vCriterio, "Contratos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOFINAL").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub CONTRATOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND EMPRESA = " + Str(CurrentCompany)
  vCriterio = vCriterio + "AND TABTIPOCONTRATO = 1"
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, 2, vCampos, vCriterio, "Contratos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOINICIAL").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If VerificaSeProcessada Then
	  	bsShowMessage("Inclusão não permitida. Rotina já processada.","E")
	  	CanContinue = False
	  End If
End Sub
