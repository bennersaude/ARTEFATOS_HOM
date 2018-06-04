'HASH: 8BC287F70921DC1693737E8F23678129
'#uses "*bsShowMessage"

Sub MontaWhere
	Dim Query As String
	Dim Query1 As String
	Dim q1 As Object
	Set q1 = NewQuery

	q1.Clear
	q1.Add("SELECT X.HANDLE                                                                   ")
	q1.Add("FROM (                                                                            ")
	q1.Add(" SELECT TGE.HANDLE                                                                ")
	q1.Add("   FROM SAM_TGE TGE,                                                              ")
	q1.Add("        SAM_ESPECIALIDADEGRUPO_EXEC EXE,                                          ")
	q1.Add("        SAM_ESPECIALIDADEGRUPO GRP,                                               ")
	q1.Add("        SAM_ESPECIALIDADE ESP,                                                    ")
	q1.Add("        SAM_PRESTADOR PRE,                                                        ")
	q1.Add("        SAM_PRESTADOR_ESPECIALIDADE PES                                           ")
	q1.Add("  WHERE PRE.HANDLE             = :PRESTADOR                                       ")
	q1.Add("    AND PES.ESPECIALIDADE      = ESP.HANDLE                                       ")
	q1.Add("    AND PES.PRESTADOR          = PRE.HANDLE                                       ")
	q1.Add("    AND EXE.ESPECIALIDADEGRUPO = GRP.HANDLE                                       ")
	q1.Add("    AND GRP.ESPECIALIDADE      = ESP.HANDLE                                       ")
	q1.Add("    AND EXE.ESPECIALIDADE      = ESP.HANDLE                                       ")
	q1.Add("    AND EXE.EVENTO             = TGE.HANDLE                                       ")
	q1.Add("UNION                                                                             ")
	q1.Add(" SELECT TGE.HANDLE                                                                ")
	q1.Add("   FROM SAM_TGE TGE,                                                              ")
	q1.Add("        SAM_PRESTADOR_REGRA REG,                                                  ")
	q1.Add("        SAM_PRESTADOR PRE                                                         ")
	q1.Add("  WHERE REG.EVENTO        = TGE.HANDLE                                            ")
	q1.Add("    AND REG.PRESTADOR     = PRE.HANDLE                                            ")
	q1.Add("    AND REG.REGRAEXCECAO  = 'R'                                                   ")
	q1.Add("    AND PRE.HANDLE        = :PRESTADOR) X                                         ")
	q1.Add("GROUP BY X.HANDLE                                                                 ")

	q1.ParamByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
	q1.Active = True

	If Not (q1.FieldByName("HANDLE").IsNull) Then
		While Not (q1.EOF)
			Query = Query + Str(q1.FieldByName("HANDLE").AsInteger) + ","
			q1.Next
		Wend

		Query1 = Mid(Query, 1, Len(Query) -1 )

		If VisibleMode Then
			EVENTO.LocalWhere = " HANDLE IN ( " + Query1 + " ) "
		Else
			EVENTO.WebLocalWhere = " HANDLE IN ( " + Query1 + " ) "
		End If

		q1.Active = False
	Else
		bsShowMessage( "Não existe especialidade ou regra cadastrada para este prestador","I")
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebVisionCode = "V_SAM_PRESTADOR_EVENTOCONSULTA" Then
			PRESTADOR.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	MontaWhere
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	MontaWhere
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim q1 As Object
	Set q1 = NewQuery

	q1.Clear

	q1.Add (" SELECT EVENTO FROM SAM_PRESTADOR_EVENTOCONSULTA WHERE PRESTADOR = :PRESTADOR")
	q1.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
	q1.Active = True

	While Not (q1.EOF)
		If q1.FieldByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger Then
			bsShowMessage("Evento já cadastrado", "E")
			CanContinue = False
		End If

		q1.Next
	Wend

	q1.Active = False
End Sub
