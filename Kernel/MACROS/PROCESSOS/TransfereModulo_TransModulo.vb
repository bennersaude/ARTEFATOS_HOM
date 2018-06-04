'HASH: 541A870D01821C195D3B244AC8869A3E

Public Sub Main

	On Error GoTo erro:

	Dim plModDestino As Long
	Dim plMotCanc As Long
	Dim psCodigoTabPrc As String
	Dim psDataAdesao As String
	Dim psCarenciaDestino As String
	Dim pbNTemCarencia As Boolean
	Dim psMensagem As String
	Dim vlHModBeneficiario As Long

	plModDestino = CLng( ServiceVar("plModDestino") )
	plMotCanc = CLng( ServiceVar("plMotCanc") )
	psCodigoTabPrc = CStr( ServiceVar("psCodigoTabPrc") )
	psDataAdesao = CStr( ServiceVar("psDataAdesao") )
	psCarenciaDestino = CStr( ServiceVar("psCarenciaDestino") )
	pbNTemCarencia = CBool( ServiceVar("pbNTemCarencia") )
	psMensagem = CStr( ServiceVar("psMensagem") )
	vlHModBeneficiario = CLng( SessionVar("HMODBENEFICIARIO") )

	If Not ( psCarenciaDestino = "" ) Then
		psCarenciaDestino = Replace( Replace( psCarenciaDestino, "&lt", ">" ), "&gt", "<")
	End If

	If plModDestino = 0 Then
			psMensagem =  "Selecione um módulo de destino."
			ServiceVar( "psMensagem" ) = psMensagem
			Exit Sub
	End If

	If plMotCanc = 0  Then
		psMensagem =  "Selecione um motivo de cancelamento"
		ServiceVar( "psMensagem" ) = psMensagem
		Exit Sub
	End If

	If psDataAdesao = "" Then
		psMensagem = "Informe a data de adesão do módulo de destino!"
		ServiceVar( "psMensagem" ) = psMensagem
		Exit Sub
	End If


    Dim qBen As BPesquisa
	Dim qBenMod As Object

	Set qBen = NewQuery
	Set qBenMod = NewQuery

  	qBen.Add("SELECT BENEFICIARIO, MODULO FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = (:HANDLE)")
  	qBen.ParamByName("HANDLE").AsInteger = vlHModBeneficiario
  	qBen.Active = True

  	qBenMod.Add("SELECT DATAADESAO                  ")
  	qBenMod.Add("  FROM SAM_BENEFICIARIO_MOD        ")
  	qBenMod.Add(" WHERE BENEFICIARIO = :BENEFICIARIO")
  	qBenMod.Add("   AND MODULO = :MODULO")
  	qBenMod.Add("   AND DATACANCELAMENTO IS NULL    ")
  	qBenMod.ParamByName("BENEFICIARIO").AsInteger = qBen.FieldByName("BENEFICIARIO").AsInteger
  	qBenMod.ParamByName("MODULO").AsInteger = qBen.FieldByName("MODULO").AsInteger

  	qBenMod.Active = True

  	If qBenMod.FieldByName("DATAADESAO").AsDateTime >=  CDate(Format(psDataAdesao, "dd/mm/yyyy")) Then
		psMensagem = "A data de adesão do módulo de destino deve ser maior que a do módulo de origem!"
		ServiceVar("psMensagem") = psMensagem
		Exit Sub
  	End If



    Dim Obj As Object

    Set Obj = CreateBennerObject("BSBEN002.Modulo")

	Dim SQL As BPesquisa
	Set SQL = NewQuery

	SQL.Add("SELECT CM.PLANO, CM.MODULO								  ")
  	SQL.Add("  FROM SAM_CONTRATO_MOD	CM 							  ")
  	SQL.Add("  JOIN SAM_BENEFICIARIO_MOD BM ON (CM.HANDLE = BM.MODULO)")
 	SQL.Add(" WHERE BM.HANDLE = (:HANDLE)							  ")

 	SQL.ParamByName("HANDLE").AsInteger = vlHModBeneficiario
 	SQL.Active = True


	Dim SQL2 As Object

	Set SQL2 = NewQuery

    SQL2.Add("SELECT PLANO , MODULO  ")
	SQL2.Add("  FROM SAM_CONTRATO_MOD")
	SQL2.Add(" WHERE HANDLE = :HANDLE")

	SQL2.ParamByName("HANDLE").AsInteger = plModDestino

	SQL2.Active = True


 	Dim vlPlanoDestino As Long
 	Dim vlModOrigem As Integer
 	Dim	vlModDestino As Integer


 	vlPlanoDestino = SQL2.FieldByName("PLANO").AsInteger
	vlModDestino = SQL2.FieldByName("MODULO").AsInteger
 	vlModOrigem = SQL.FieldByName("MODULO").AsInteger



   psMensagem = Obj.TransfereModulo(CurrentSystem, _
									vlModDestino,  _
									vlPlanoDestino, _
									vlHModBeneficiario, _
									vlModOrigem, _
									plMotCanc, _
									psCodigoTabPrc, _
									CDate(Format(psDataAdesao, "dd/mm/yyyy")), _
									psCarenciaDestino, _
									pbNTemCarencia )

	ServiceVar("psMensagem") = psMensagem
	Exit Sub

	erro:
		ServiceVar("psMensagem") = Err.Description


End Sub
