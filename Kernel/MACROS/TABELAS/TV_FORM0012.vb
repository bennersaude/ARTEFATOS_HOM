'HASH: CAF2C0D03E25499E0EB832EEAF1AE70E
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim vDataFinalSuspensao As Date


  	Dim BSBen001 As Object
	Dim cancelamentoDLL As Object

	Set cancelamentoDLL = CreateBennerObject("SAMCANCELAMENTO.Reativar")
  	Set BSBen001 = CreateBennerObject("BSBen001.Beneficiario")


	Dim QFAMILIA As Object
	Set QFAMILIA = NewQuery

	QFAMILIA.Add("SELECT CONTRATO,FAMILIA ")
	QFAMILIA.Add("	FROM SAM_FAMILIA_MOD")
	QFAMILIA.Add(" WHERE HANDLE = :HANDLE")


	QFAMILIA.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_FAMILIA_MOD")
	QFAMILIA.Active = True

	If BSBen001.VerificaSuspensao(CurrentSystem, _
	                                   	0, _
                                    	QFAMILIA.FieldByName("FAMILIA").AsInteger, _
                                    	QFAMILIA.FieldByName("CONTRATO").AsInteger, _
                                    	vDataFinalSuspensao)Then
    		bsShowMessage("Não é permitido reativar o módulo por motivo de suspensão!", "I")
    		CanContinue = False
			Set BSBen001 = Nothing
    		Exit Sub
  	End If


	' SMS 94182 - Paulo Melo - 05/03/2008 - A bloqueio para a data de reativação não ser nula e nem menor que a data de adesão do módulo da família.
  	Dim qDataAdesao As Object
  	Set qDataAdesao = NewQuery

	qDataAdesao.Add("SELECT DATAADESAO")
  	qDataAdesao.Add("FROM SAM_FAMILIA_MOD")
  	qDataAdesao.Add("WHERE HANDLE = :HFAMILIAMOD")
  	qDataAdesao.ParamByName("HFAMILIAMOD").AsInteger = RecordHandleOfTable("SAM_FAMILIA_MOD")
  	qDataAdesao.Active = True

	If CurrentQuery.FieldByName("DATAREATIVACAO").AsString = "" Then
		bsShowMessage("Data de reativação não pode ser nula.", "E")
		CanContinue = False
		CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
		Exit Sub
	End If
  	If (CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime < qDataAdesao.FieldByName("DATAADESAO").AsDateTime) Then
    	bsShowMessage("Data de reativação inferior à data de adesão do beneficiário: " + CStr(qDataAdesao.FieldByName("DATAADESAO").AsDateTime), "E")
    	CanContinue = False
    	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
    	Exit Sub
	End If
  ' SMS 94182 - Paulo Melo - FIM

	bsShowMessage(cancelamentoDLL.FamiliaModulo(CurrentSystem, RecordHandleOfTable("SAM_FAMILIA_MOD"), CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime), "I")



  	Set BSBen001Dll = Nothing
	Set QFAMILIA = Nothing
	Set cancelamentoDLL = Nothing

End Sub
