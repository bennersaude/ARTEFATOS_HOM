'HASH: F698ABA86F3627803626B2C570BA7025
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim vDataFinalSuspensao As Date
	Dim Obj As Object
	Dim BSBen001Dll As Object


  	Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")

  	If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    0, _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
	   		bsShowMessage("Não é permitido reativar o contrato por motivo de suspensão!", "E")
		    Set BSBen001Dll = Nothing
	   		CanContinue = False
		    Exit Sub
  	End If
  	Set BSBen001Dll = Nothing


	Set Obj = CreateBennerObject("SAMCANCELAMENTO.Reativar")

	' SMS 94182 - Paulo Melo - 05/03/2008 - bloqueio para a data de reativação não ser nula e nem menor que a data de adesão do contrato.
  	Dim qDataAdesao As Object
  	Set qDataAdesao = NewQuery

	qDataAdesao.Add("SELECT DATAADESAO")
  	qDataAdesao.Add("FROM SAM_CONTRATO")
  	qDataAdesao.Add("WHERE HANDLE = :HCONTRATO")
  	qDataAdesao.ParamByName("HCONTRATO").AsInteger = RecordHandleOfTable("SAM_CONTRATO")
  	qDataAdesao.Active = True

	If CurrentQuery.FieldByName("DATAREATIVACAO").AsString = "" Then
		bsShowMessage("Data de reativação não pode ser nula.", "E")
		CanContinue = False
		CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
		Exit Sub
	End If
  	If (CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime < qDataAdesao.FieldByName("DATAADESAO").AsDateTime) Then
    	bsShowMessage("Data de reativação inferior à data de adesão do contrato: " + CStr(qDataAdesao.FieldByName("DATAADESAO").AsDateTime), "E")
    	CanContinue = False
    	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
    	Exit Sub
    End If
  ' SMS 94182 - Paulo Melo - FIM

    If VisibleMode Then
	 	bsShowMessage(Obj.Contrato(CurrentSystem, RecordHandleOfTable("SAM_CONTRATO")   ,CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime), "I")
 	ElseIf WebMode Then

		Dim SQL As Object
	    Set SQL = NewQuery
	    SQL.Add("SELECT DATACANCELAMENTO ")
	    SQL.Add("  FROM SAM_CONTRATO ")
    	SQL.Add(" WHERE HANDLE = :HANDLE")
    	SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_CONTRATO")

	    SQL.Active = True

	    If SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    	  bsShowMessage("O Contrato não está cancelado !", "E")
    	  Set SQL = Nothing
	      CanContinue = False
	      Exit Sub
    	End If

	    Set SQL = Nothing

	 	bsShowMessage(Obj.Contrato(CurrentSystem, RecordHandleOfTable("SAM_CONTRATO") ,CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime), "I")
 	End If

	 Set Obj = Nothing

End Sub
