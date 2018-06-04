'HASH: 37A9491B4468B73D80A09DB782480D83
'MACRO TABELA: SAM_MEMBRO_ESPECIALIDADEGRUPO
'#Uses "*bsShowMessage"

Dim vCondicao As String

Public Sub ESPECIALIDADEGRUPO_OnPopup(ShowPopup As Boolean)
	UpdateLastUpdate("SAM_ESPECIALIDADEGRUPO")

	Dim SQL1 As Object
	Dim SQL2 As Object
	Dim SQL3 As Object
	Dim SQL4 As Object
	Dim SQL5 As Object
	Dim ENTIDADE As Object
	Set SQL1 = NewQuery

	SQL1.Add("SELECT * FROM SAM_MEMBRO_ESPECIALIDADE WHERE HANDLE=:HANDLE")

	SQL1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MEMBROESPECIALIDADE").Value
	SQL1.Active = True

	Set ENTIDADE = NewQuery

	ENTIDADE.Add("SELECT ENTIDADE, PRESTADOR FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE HANDLE=:HANDLE")

	ENTIDADE.ParamByName("HANDLE").Value = SQL1.FieldByName("CORPOCLINICO").Value
	ENTIDADE.Active = True

	Set SQL2 = NewQuery

	SQL2.Add("SELECT A.* ")
	SQL2.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE A                                         ")
	SQL2.Add(" WHERE A.PRESTADOR = :ENTIDADE                                               ")
	SQL2.Add("   AND A.ESPECIALIDADE = :ESPECIALIDADE                                      ")
	SQL2.Add("   AND A.DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)  ")
	SQL2.Add("   AND A.ESPECIALIDADE IN (SELECT B.ESPECIALIDADE                            ")
	SQL2.Add("                 FROM SAM_PRESTADOR_ESPECIALIDADE  B                         ")
	SQL2.Add("              WHERE B.DATAINICIAL <= :DATA AND (B.DATAFINAL >= :DATA OR B.DATAFINAL IS NULL)")
	SQL2.Add("                  AND B.ESPECIALIDADE = A.ESPECIALIDADE                      ")
	SQL2.Add("                  AND B.PRESTADOR = :PRESTADOR)                              ")

	SQL2.ParamByName("ENTIDADE").Value = ENTIDADE.FieldByName("ENTIDADE").AsInteger
	SQL2.ParamByName("PRESTADOR").Value = ENTIDADE.FieldByName("PRESTADOR").AsInteger
	SQL2.ParamByName("ESPECIALIDADE").Value = SQL1.FieldByName("ESPECIALIDADE").AsInteger
	SQL2.ParamByName("DATA").Value = ServerDate
	SQL2.Active = True


	Set SQL3 = NewQuery

	SQL3.Add("SELECT A.* ")
	SQL3.Add("  FROM SAM_PRESTADOR_ESPECIALIDADEGRP A                                                     ")
	SQL3.Add(" WHERE A.PRESTADOR     = :ENTIDADE                                                          ")
	SQL3.Add("   AND A.ESPECIALIDADE = :ESPECIALIDADE                                                     ")
	SQL3.Add("   AND A.DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)                 ")
	SQL3.Add("   AND A.ESPECIALIDADE IN (SELECT B.ESPECIALIDADE                                           ")
	SQL3.Add("                 FROM SAM_PRESTADOR_ESPECIALIDADEGRP  B                                     ")
	SQL3.Add("                WHERE B.DATAINICIAL <= :DATA AND (B.DATAFINAL >= :DATA OR B.DATAFINAL IS NULL)")
	SQL3.Add("                  AND B.ESPECIALIDADE = A.ESPECIALIDADE                                     ")
	SQL3.Add("                  AND B.ESPECIALIDADEGRUPO = A.ESPECIALIDADEGRUPO                           ")
	SQL3.Add("                  AND B.PRESTADOR = :PRESTADOR)                                             ")

	SQL3.ParamByName("ENTIDADE").Value = ENTIDADE.FieldByName("ENTIDADE").AsInteger
	SQL3.ParamByName("PRESTADOR").Value = ENTIDADE.FieldByName("PRESTADOR").AsInteger
	SQL3.ParamByName("ESPECIALIDADE").Value = SQL1.FieldByName("ESPECIALIDADE").Value
	SQL3.ParamByName("DATA").Value = ServerDate
	SQL3.Active = True

	If Not SQL3.EOF Then
		vCondicao = ""
		vCondicao = vCondicao + "SAM_ESPECIALIDADEGRUPO.HANDLE "
		vCondicao = vCondicao + "IN (SELECT G.ESPECIALIDADEGRUPO FROM SAM_PRESTADOR_ESPECIALIDADEGRP G WHERE G.HANDLE = " + SQL3.FieldByName("HANDLE").AsInteger

		SQL3.Next

		While Not SQL3.EOF
			vCondicao = vCondicao + "       OR G.HANDLE      = " + SQL3.FieldByName("HANDLE").AsInteger
			SQL3.Next
		Wend

		vCondicao = vCondicao + "  )"
	Else
		If Not SQL2.EOF Then
			Set SQL4 = NewQuery

			SQL4.Add("SELECT A.* ")
			SQL4.Add("  FROM SAM_PRESTADOR_ESPECIALIDADEGRP A                                     ")
			SQL4.Add(" WHERE A.PRESTADOR     = :ENTIDADE                                          ")
			SQL4.Add("   AND A.ESPECIALIDADE = :ESPECIALIDADE                                     ")
			SQL4.Add("   AND A.DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL) ")

			SQL4.ParamByName("ENTIDADE").Value = ENTIDADE.FieldByName("ENTIDADE").AsInteger
			SQL4.ParamByName("ESPECIALIDADE").Value = SQL2.FieldByName("ESPECIALIDADE").Value
			SQL4.ParamByName("DATA").Value = ServerDate
			SQL4.Active = True

			Set SQL5 = NewQuery

			SQL5.Add("SELECT A.* ")
			SQL5.Add("  FROM SAM_PRESTADOR_ESPECIALIDADEGRP A                                     ")
			SQL5.Add(" WHERE A.PRESTADOR     = :PRESTADOR                                         ")
			SQL5.Add("   AND A.ESPECIALIDADE = :ESPECIALIDADE                                     ")
			SQL5.Add("   AND A.DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL) ")

			SQL5.ParamByName("PRESTADOR").Value = ENTIDADE.FieldByName("PRESTADOR").AsInteger
			SQL5.ParamByName("ESPECIALIDADE").Value = SQL2.FieldByName("ESPECIALIDADE").Value
			SQL5.ParamByName("DATA").Value = ServerDate
			SQL5.Active = True

			If SQL4.EOF And SQL5.EOF Then
				vCondicao = ""
				vCondicao = vCondicao + "SAM_ESPECIALIDADEGRUPO.HANDLE "
				vCondicao = vCondicao + " IN (SELECT HANDLE FROM SAM_ESPECIALIDADEGRUPO WHERE ESPECIALIDADE = " + SQL2.FieldByName("ESPECIALIDADE").Value + ")"
			Else
				vCondicao = "SAM_ESPECIALIDADEGRUPO.HANDLE IS NULL"
			End If
		Else
			vCondicao = "SAM_ESPECIALIDADEGRUPO.HANDLE IS NULL"
		End If
	End If

	ESPECIALIDADEGRUPO.LocalWhere = vCondicao

	Set SQL1 = Nothing
	Set SQL2 = Nothing
	Set SQL3 = Nothing
	Set SQL4 = Nothing
	Set SQL5 = Nothing
	Set ENTIDADE = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	If Not VisibleMode Then
		ESPECIALIDADEGRUPO_OnPopup(False)
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("PERMITEEXECUTAR").AsString <> "S" And _
	   CurrentQuery.FieldByName("PERMITERECEBER").AsString <> "S" Then
		CanContinue = False
		bsShowMessage("Deve selecionar Permite Executar e/ou Permite Receber", "E")
		Exit Sub
	End If
End Sub
