'HASH: CAA426A44432E86575FD03E362E078C3
 
'#Uses "*bsShowMessage"
Public Sub BOTAOMOVFIN_OnClick()
	Dim qAux As Object

	Set qAux = NewQuery

	qAux.Add("SELECT HANDLE FROM R_RELATORIOS")
	qAux.Add(" WHERE CODIGO = 'BEN070'")
 	qAux.Active = True

	ReportPreview(qAux.FieldByName("HANDLE").AsInteger, "", False, False)

 	Set qAux = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  	Dim vDataFinalSuspensao As Date
  	Dim BSBen001Dll As Object
  	Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")

 	Dim SQL As Object
 	Set SQL = NewQuery

 	SQL.Add("SELECT DATACANCELAMENTO")
 	SQL.Add("  FROM SAM_CONTRATO_MOD")
 	SQL.Add(" WHERE HANDLE = :HANDLE")

 	SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_CONTRATO_MOD")

 	SQL.Active = True

	If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
               	                    	0, _
                                   		0, _
                               			RecordHandleOfTable("SAM_CONTRATO"), _
                               			vDataFinalSuspensao)Then
		bsShowMessage("Não é permitido cancelar o módulo por motivo de suspensão!", "E")
		CanContinue = False
		Exit Sub
  	End If


	Dim Obj As Object
    Set Obj = CreateBennerObject("SAMCANCELAMENTO.Cancelar")

	If VisibleMode Then
	  	bsShowMessage(Obj.ContratoModulo(CurrentSystem, RecordHandleOfTable("SAM_CONTRATO_MOD"),CurrentQuery.FieldByName("DATACANC").AsDateTime, CurrentQuery.FieldByName("MOTIVO").AsInteger), "I")
	ElseIf WebMode Then
		  	If SQL.FieldByName("DATACANCELAMENTO").IsNull Then
				bsShowMessage(Obj.ContratoModulo(CurrentSystem, RecordHandleOfTable("SAM_CONTRATO_MOD"),CurrentQuery.FieldByName("DATACANC").AsDateTime, CurrentQuery.FieldByName("MOTIVO").AsInteger), "I")
		  	Else
    			bsShowMessage("Módulo do contrato não cancelado!", "E")
    			CanContinue = False
    			Exit Sub
		  	End If

			Set SQL = Nothing
	End If

	Set Obj = Nothing
	Set BSBen001Dll = Nothing
End Sub
