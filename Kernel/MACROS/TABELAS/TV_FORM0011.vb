'HASH: 0A27FE53ABBC0AF57F286E4A887020BF
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim SQL As BPesquisa
	Dim Obj As Object

	'Confirmar a reativação
    If bsShowMessage("Confirma a reativação da Família ?", "Q") = vbYes Then

	  Dim Interface As Object
	  Set Interface = CreateBennerObject("SAMCANCELAMENTO.Reativar")


   	' SMS 94182 - Paulo Melo - 05/03/2008 - A bloqueio para a data de reativação não ser nula e nem menor que a data de adesão da família.
  	  Dim qDataAdesao As Object
  	  Set qDataAdesao = NewQuery

	  Dim Familia As Long
  	  If RecordHandleOfTable("SAM_FAMILIA")<=0 Then
		  Familia = CLng(SessionVar("HFAMILIA_REATIVACAO"))
  	  Else
		  Familia = RecordHandleOfTable("SAM_FAMILIA")
  	  End If

	  qDataAdesao.Add("SELECT DATAADESAO")
  	  qDataAdesao.Add("FROM SAM_FAMILIA")
  	  qDataAdesao.Add("WHERE HANDLE = :HFAMILIA")
  	  qDataAdesao.ParamByName("HFAMILIA").AsInteger = Familia
  	  qDataAdesao.Active = True

	  If CurrentQuery.FieldByName("DATAREATIVACAO").AsString = "" Then
		bsShowMessage("Data de reativação não pode ser nula.", "E")
		CanContinue = False
		CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
		Exit Sub
	  End If
  	  If (CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime < qDataAdesao.FieldByName("DATAADESAO").AsDateTime) Then
    	bsShowMessage("Data de reativação inferior à data de adesão da família: " + CStr(qDataAdesao.FieldByName("DATAADESAO").AsDateTime), "E")
    	CanContinue = False
    	CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime = ServerDate
    	Exit Sub
	  End If
    ' SMS 94182 - Paulo Melo - FIM

      If VisibleMode Then
	 	bsShowMessage(Interface.Familia(CurrentSystem, Familia  ,CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime), "I")
 	  ElseIf WebMode Then

	    Set SQL = NewQuery

	    SQL.Add("SELECT DATACANCELAMENTO, MOTIVOCANCELAMENTO, CONTRATO ")
	    SQL.Add("  FROM SAM_FAMILIA ")
    	SQL.Add(" WHERE HANDLE = :HANDLE")
    	SQL.ParamByName("HANDLE").AsInteger = Familia

	    SQL.Active = True

	    If SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    	  bsShowMessage("A Família não está cancelada !", "E")
    	  Set SQL = Nothing
	      CanContinue = False
	      Exit Sub
    	End If

	  	'Se o titular estiver falecido,impedir reativar família
	  	Dim QryFalecimentoTitular As Object
	  	Set QryFalecimentoTitular = NewQuery
	  	QryFalecimentoTitular.Add("SELECT MOTIVOFALECIMENTOTITULAR, MOTIVOFALECIMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  		QryFalecimentoTitular.Active = True
  		If SQL.FieldByName("MOTIVOCANCELAMENTO").Value = QryFalecimentoTitular.FieldByName("MOTIVOFALECIMENTOTITULAR").Value Then
	    	bsShowMessage("Necessário reativar tilular - Titular falecido !", "E")
	    	Set SQL = Nothing
	    	CanContinue = False
	    	Exit Sub
  		End If

		' Se o CONTRATO estiver cancelado não pode
	  	Dim SQL2 As Object
	  	Set SQL2 = NewQuery
	  	SQL2.Add("SELECT DATACANCELAMENTO FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
	  	SQL2.ParamByName("CONTRATO").AsInteger = SQL.FieldByName("CONTRATO").AsInteger
	  	SQL2.Active = True
	  	If Not SQL2.FieldByName("DATACANCELAMENTO").IsNull Then
	    	SQL2.Active = False
	    	Set SQL2 = Nothing
	    	bsShowMessage("Não é permitido reativar famílias nesse contrato - Contrato está Cancelado !", "E")
	    	CanContinue = False
	    	Exit Sub
  		End If

		Set SQL2 = Nothing
   		Set SQL = Nothing


	 	bsShowMessage(Interface.Familia(CurrentSystem, Familia ,CurrentQuery.FieldByName("DATAREATIVACAO").AsDateTime), "I")
 	  End If
 	End If

	Set Obj = Nothing

End Sub
