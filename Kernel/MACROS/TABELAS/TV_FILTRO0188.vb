'HASH: 6EACAD64DD350BC5DDD9883CC458A372
 
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("TV_FILTRO0188") <> "" Then
    	XMLToDataset(UserVar("TV_FILTRO0188"),CurrentQuery.TQuery)
  	End If
End Sub

Public Sub TABLE_AfterPost()
	If CurrentQuery.FieldByName("GRAVAPARAMETROS").AsBoolean = True Then
		UserVar("TV_FILTRO0188") = DatasetToXML(CurrentQuery.TQuery,"")
	Else
		UserVar("TV_FILTRO0188") = ""
	End If
End Sub

Public Sub TABLE_AfterScroll()

	Dim qParamGeralAtendimento As Object
	Dim diasConsultaLiberadas As Integer

	Set qParamGeralAtendimento = NewQuery
	qParamGeralAtendimento.Active = False
	qParamGeralAtendimento.Add("SELECT DIASCONSULTALIBERADAS FROM SAM_PARAMETROSATENDIMENTO")
	qParamGeralAtendimento.Active = True

	If CurrentQuery.FieldByName("GRAVAPARAMETROS").AsBoolean = False Then
		CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = ServerDate - qParamGeralAtendimento.FieldByName("DIASCONSULTALIBERADAS").AsInteger
		CurrentQuery.FieldByName("DATAFINAL").AsDateTime = ServerDate
		CurrentQuery.FieldByName("FILIAL").AsInteger = CurrentBranch
	End If

	Set qParamGeralAtendimento = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If (CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) Then
      bsShowMessage("A data inicial não pode ser maior que a data final!", "E")
	  CanContinue = False
	  Exit Sub
    End If

    If CurrentQuery.FieldByName("TABTIPOFILTRO").AsInteger = 1 Then
		CurrentQuery.FieldByName("PROTOCOLOATENDIMENTO").Clear
	Else
		CurrentQuery.FieldByName("DATAINICIAL").Clear
		CurrentQuery.FieldByName("DATAFINAL").Clear
	End If

End Sub
