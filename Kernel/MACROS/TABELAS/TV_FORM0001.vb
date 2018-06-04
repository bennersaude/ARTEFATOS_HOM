'HASH: 616C5B80EDD7BE02DDC93385EBEE20F1
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATA").AsDateTime = ServerDate
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Obj As Object
	Set Obj = CreateBennerObject("BSBEN009.Beneficiario")
	If VisibleMode Then
    	bsShowMessage(Obj.CancelaCartaoIndividual(CurrentSystem, CLng(SessionVar("HANDLE")),CurrentQuery.FieldByName("DATA").AsDateTime, CurrentQuery.FieldByName("MOTIVO").AsInteger), "I")
    ElseIf WebMode Then

      Dim SQL As Object
	  Set SQL = NewQuery

	  	SQL.Add("SELECT SITUACAO ")
	  	SQL.Add("  FROM SAM_BENEFICIARIO_CARTAOIDENTIF ")
	  	SQL.Add(" WHERE HANDLE = :HANDLE")
  		SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO_CARTAOIDENTIF")

        SQL.Active = True

	   If SQL.FieldByName("SITUACAO").AsString = "C" Then
	   		bsShowMessage("Cartão já cancelado", "E")
	   		CanContinue = False
	   End If
       bsShowMessage(Obj.CancelaCartaoIndividual(CurrentSystem, RecordHandleOfTable("SAM_BENEFICIARIO_CARTAOIDENTIF"),CurrentQuery.FieldByName("DATA").AsDateTime, CurrentQuery.FieldByName("MOTIVO").AsInteger), "I")

       Set SQL = Nothing
    End If
	Set Obj = Nothing

End Sub
