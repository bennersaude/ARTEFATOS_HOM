'HASH: EC93732D711B65C5A12C18186629481E
 '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim qData As Object
	Set qData = NewQuery

	qData.Add("SELECT DATARETIRADA				   ")
	qData.Add("FROM SAM_BENEFICIARIO_CARTAOIDENTIF ")
	qData.Add("WHERE HANDLE = :HANDLE       	   ")
	If VisibleMode Then
		qData.ParamByName("HANDLE").AsString = SessionVar("HANDLE")
	ElseIf WebMode Then
		qData.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO_CARTAOIDENTIF")
	End If
	qData.Active = True

	Dim qAtualiza As Object

	If Not qData.FieldByName("DATARETIRADA").IsNull Then
    	bsShowMessage("A retirada do cartão já foi registrada no sistema!", "E")
		CanContinue = False
    	Exit Sub
  	End If

	If  CurrentQuery.FieldByName("DATARETIRADA").AsDateTime  > ServerDate Then
    	bsShowMessage("Não permitir data de retirada maior que a data atual", "E")
		CanContinue = False
      	Set qAtualiza = Nothing
      	Exit Sub
    End If

	If Not InTransaction Then StartTransaction
		Set qAtualiza = NewQuery
    	qAtualiza.Active = False
    	qAtualiza.Clear
    	qAtualiza.Add("UPDATE SAM_BENEFICIARIO_CARTAOIDENTIF ")
    	qAtualiza.Add("   SET DATARETIRADA  = :DATA,         ")
    	qAtualiza.Add("       NOMERETIRADA  = :RESPONSAVEL,  ")
    	qAtualiza.Add("       RGCPFRETIRADA = :RGCPF         ")
    	qAtualiza.Add(" WHERE HANDLE = :HANDLE               ")
    	qAtualiza.ParamByName("DATA").AsDateTime      = CurrentQuery.FieldByName("DATARETIRADA").AsDateTime
    	qAtualiza.ParamByName("RESPONSAVEL").AsString = CurrentQuery.FieldByName("NOMERETIRADA").AsString
    	qAtualiza.ParamByName("RGCPF").AsString       = CurrentQuery.FieldByName("RGCPFRETIRADA").AsString
		If VisibleMode Then
    		qAtualiza.ParamByName("HANDLE").AsString = SessionVar("HANDLE")
    	ElseIf WebMode Then
    		qAtualiza.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO_CARTAOIDENTIF")
		End If

    	qAtualiza.ExecSQL
		Set qAtualiza = Nothing
    If InTransaction Then Commit
    Set qData = Nothing

  '  If InTransaction Then Rollback

End Sub
