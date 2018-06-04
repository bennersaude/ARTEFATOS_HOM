'HASH: A6D2A9AA0A1B0FE6EF026917F26357B2
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  	Dim samCancelamentoDLL As Object
	Dim QDADOS As Object
    Dim BSBEN001 As Object
    Dim vDatafinalSuspensao As Date


	Set QDADOS = NewQuery
	QDADOS.Add("SELECT CONTRATO, DATACANCELAMENTO")
   	QDADOS.Add("  FROM SAM_CONTRATO_MOD")
   	QDADOS.Add(" WHERE HANDLE = :HANDLE")

    Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
 	Set samCancelamentoDLL = CreateBennerObject("SAMCANCELAMENTO.Reativar")

	QDADOS.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_CONTRATO_MOD")

   	QDADOS.Active = True


	If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
	                                   0, _
    	                               0, _
        	                           QDADOS.FieldByName("CONTRATO").AsInteger, _
            	                       vDatafinalSuspensao)Then
	    	bsShowMessage("Não é permitido reativar o módulo por motivo de suspensão!", "E")
	    	CanContinue = False
	    	Set BSBen001Dll = Nothing
		    Exit Sub
  	End If
  	Set BSBen001Dll = Nothing

  	' SMS 94182 - Paulo Melo - 05/03/2008 - A bloqueio para a data de reativação não ser nula e nem menor que a data de adesão do módulo do contrato.
  	Dim qDataAdesao As Object
  	Set qDataAdesao = NewQuery

	qDataAdesao.Add("SELECT DATAADESAO")
  	qDataAdesao.Add("FROM SAM_CONTRATO_MOD")
  	qDataAdesao.Add("WHERE HANDLE = :HCONTRATOMOD")
  	qDataAdesao.ParamByName("HCONTRATOMOD").AsInteger = RecordHandleOfTable("SAM_CONTRATO_MOD")
  	qDataAdesao.Active = True

	If CurrentQuery.FieldByName("DATAREATIVACAO").AsString = "" Then
		bsShowMessage("Data de reativação não pode ser nula.", "E")
		CanContinue = False
		CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
		Exit Sub
	End If
  	If (CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime < qDataAdesao.FieldByName("DATAADESAO").AsDateTime) Then
    	bsShowMessage("Data de reativação inferior à data de adesão do módulo do contrato: " + CStr(qDataAdesao.FieldByName("DATAADESAO").AsDateTime), "E")
    	CanContinue = False
    	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
    	Exit Sub
	End If
  ' SMS 94182 - Paulo Melo - FIM

	If VisibleMode Then
		bsShowMessage(samCancelamentoDLL.ContratoModulo(CurrentSystem, RecordHandleOfTable("SAM_CONTRATO_MOD"), CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime), "I")
	ElseIf WebMode Then
		If Not QDADOS.FieldByName("DATACANCELAMENTO").IsNull Then
			Dim QCONTRATO As Object
      		Set QCONTRATO = NewQuery
   			QCONTRATO.Add("SELECT TABTIPOCONTRATO, DATACANCELAMENTO FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
      		QCONTRATO.ParamByName("CONTRATO").Value = QDADOS.FieldByName("CONTRATO").AsInteger
   			QCONTRATO.Active = True
     		If Not QCONTRATO.FieldByName("DATACANCELAMENTO").IsNull Then
	        	bsShowMessage("Não é permitido reativar módulos nesse contrato - Contrato está Cancelado!", "E")
	        	CanContinue = False
        		Exit Sub
      		End If
			bsShowMessage(samCancelamentoDLL.ContratoModulo(CurrentSystem,  RecordHandleOfTable("SAM_CONTRATO_MOD"), CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime), "I")

    	  	Set QCONTRATO = Nothing
 	  	 Else
 	   		bsShowMessage("Contrato não cancelado!","I")
         End If

    End If

    Set QDADOS = Nothing



End Sub
