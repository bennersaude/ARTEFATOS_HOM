'HASH: 5C1D4B51B16DF70E3C1F086BBE94A990
'MACRO SAM_LIVROCONFIG
'Alterada -13/01/2003 -Claudemir
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterEdit()
	vCondicao = "NOT ( "

	If VisibleMode Then
		ORDEM2.LocalWhere = vCondicao + "HANDLE = @ORDEM1)"
		ORDEM3.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2)"
		ORDEM4.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2 OR HANDLE = @ORDEM3)"
		ORDEM5.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2 OR HANDLE = @ORDEM3 OR HANDLE = @ORDEM4)"
		ORDEM6.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2 OR HANDLE = @ORDEM3 OR HANDLE = @ORDEM4 " + _
									 "OR HANDLE = @ORDEM5)"
		ORDEM7.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2 OR HANDLE = @ORDEM3 OR HANDLE = @ORDEM4 " + _
									 "OR HANDLE = @ORDEM5 OR HANDLE = @ORDEM6)"
		ORDEM8.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2 OR HANDLE = @ORDEM3 OR HANDLE = @ORDEM4 " + _
									 "OR HANDLE = @ORDEM5 OR HANDLE = @ORDEM6 OR HANDLE = @ORDEM7)"
		ORDEM9.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2 OR HANDLE = @ORDEM3 OR HANDLE = @ORDEM4 " + _
									 "OR HANDLE = @ORDEM5 OR HANDLE = @ORDEM6 OR HANDLE = @ORDEM7 OR HANDLE = @ORDEM8)"
		ORDEM10.LocalWhere = vCondicao + "HANDLE = @ORDEM1 OR HANDLE = @ORDEM2 OR HANDLE = @ORDEM3 OR HANDLE = @ORDEM4 " + _
									  "OR HANDLE = @ORDEM5 OR HANDLE = @ORDEM6 OR HANDLE = @ORDEM7 OR HANDLE = @ORDEM8 " + _
									  "OR HANDLE = @ORDEM9)"
	Else
		ORDEM2.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1))"
		ORDEM3.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2))"
		ORDEM4.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2) OR HANDLE = @CAMPO(ORDEM3))"
		ORDEM5.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2) OR HANDLE = @CAMPO(ORDEM3) " + _
										"OR HANDLE = @CAMPO(ORDEM4))"
		ORDEM6.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2) OR HANDLE = @CAMPO(ORDEM3) " + _
										"OR HANDLE = @CAMPO(ORDEM4) OR HANDLE = @CAMPO(ORDEM5))"
		ORDEM7.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2) OR HANDLE = @CAMPO(ORDEM3) " + _
										"OR HANDLE = @CAMPO(ORDEM4) OR HANDLE = @CAMPO(ORDEM5) OR HANDLE = @CAMPO(ORDEM6))"
		ORDEM8.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2) OR HANDLE = @CAMPO(ORDEM3) " + _
										"OR HANDLE = @CAMPO(ORDEM4) OR HANDLE = @CAMPO(ORDEM5) OR HANDLE = @CAMPO(ORDEM6) " + _
										"OR HANDLE = @CAMPO(ORDEM7))"
		ORDEM9.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2) OR HANDLE = @CAMPO(ORDEM3) " + _
										"OR HANDLE = @CAMPO(ORDEM4) OR HANDLE = @CAMPO(ORDEM5) OR HANDLE = @CAMPO(ORDEM6) " + _
										"OR HANDLE = @CAMPO(ORDEM7) OR HANDLE = @CAMPO(ORDEM8))"
		ORDEM10.WebLocalWhere = vCondicao + "HANDLE = @CAMPO(ORDEM1) OR HANDLE = @CAMPO(ORDEM2) OR HANDLE = @CAMPO(ORDEM3) " + _
										 "OR HANDLE = @CAMPO(ORDEM4) OR HANDLE = @CAMPO(ORDEM5) OR HANDLE = @CAMPO(ORDEM6) " + _
										 "OR HANDLE = @CAMPO(ORDEM7) OR HANDLE = @CAMPO(ORDEM8) OR HANDLE = @CAMPO(ORDEM9))"

	End If
End Sub

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATA").AsDateTime = ServerDate
	CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser

	vCondicao = ""

	If VisibleMode Then
		vCondicao = vCondicao + "SAM_LIVROQUEBRAS.HANDLE "

		ORDEM2.LocalWhere = vCondicao + "NOT IN (@ORDEM1)"
		ORDEM3.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2)"
		ORDEM4.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2, @ORDEM3)"
		ORDEM5.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2, @ORDEM3, @ORDEM4)"
		ORDEM6.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2, @ORDEM3, @ORDEM4, @ORDEM5)"
		ORDEM7.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2, @ORDEM3, @ORDEM4, @ORDEM5, @ORDEM6)"
		ORDEM8.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2, @ORDEM3, @ORDEM4, @ORDEM5, @ORDEM6, @ORDEM7)"
		ORDEM9.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2, @ORDEM3, @ORDEM4, @ORDEM5, @ORDEM6, @ORDEM7, @ORDEM8)"
		ORDEM10.LocalWhere = vCondicao + "NOT IN (@ORDEM1, @ORDEM2, @ORDEM3, @ORDEM4, @ORDEM5, @ORDEM6, @ORDEM7, @ORDEM8, @ORDEM9)"
	Else
		vCondicao = vCondicao + "A.HANDLE "

		ORDEM2.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1))"
		ORDEM3.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2))"
		ORDEM4.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2), @CAMPO(ORDEM3))"
		ORDEM5.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2), @CAMPO(ORDEM3), @CAMPO(ORDEM4))"
		ORDEM6.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2), @CAMPO(ORDEM3), @CAMPO(ORDEM4), " + _
			"@CAMPO(ORDEM5))"
		ORDEM7.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2), @CAMPO(ORDEM3), @CAMPO(ORDEM4), " + _
			"@CAMPO(ORDEM5), @CAMPO(ORDEM6))"
		ORDEM8.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2), @CAMPO(ORDEM3), @CAMPO(ORDEM4), " + _
			"@CAMPO(ORDEM5), @CAMPO(ORDEM6), @CAMPO(ORDEM7))"
		ORDEM9.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2), @CAMPO(ORDEM3), @CAMPO(ORDEM4), " + _
			"@CAMPO(ORDEM5), @CAMPO(ORDEM6), @CAMPO(ORDEM7), @CAMPO(ORDEM8))"
		ORDEM10.WebLocalWhere = vCondicao + "NOT IN (@CAMPO(ORDEM1), @CAMPO(ORDEM2), @CAMPO(ORDEM3), @CAMPO(ORDEM4), " + _
			"@CAMPO(ORDEM5), @CAMPO(ORDEM6), @CAMPO(ORDEM7), @CAMPO(ORDEM8), @CAMPO(ORDEM9))"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub ORDEM2_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM1").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo1!", "I")
	End If
End Sub

Public Sub ORDEM3_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM2").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo2!", "I")
	End If
End Sub

Public Sub ORDEM4_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM3").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo3!", "I")
	End If
End Sub

Public Sub ORDEM5_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM4").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo4!", "I")
	End If
End Sub

Public Sub ORDEM6_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM5").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo5!", "I")
	End If
End Sub

Public Sub ORDEM7_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM6").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo6!", "I")
	End If
End Sub

Public Sub ORDEM8_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM7").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo7!", "I")
	End If
End Sub

Public Sub ORDEM9_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM8").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo8!", "I")
	End If
End Sub

Public Sub ORDEM10_OnPopup(ShowPopup As Boolean)
	If CurrentQuery.FieldByName("ORDEM9").IsNull Then
		ShowPopup = False
		bsShowMessage("Informar Campo9!", "I")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	CurrentQuery.FieldByName("DATA").AsDateTime = ServerDate
	CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
