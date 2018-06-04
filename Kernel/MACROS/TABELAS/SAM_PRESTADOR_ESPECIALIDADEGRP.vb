'HASH: 3114BE76A9C8C76F3FAB4BF0BE308263
'Macro: SAM_PRESTADOR_ESPECIALIDADEGRP
'#Uses "*liberaEspecialidade"
'#Uses "*bsShowMessage"
'#Uses "*RegistrarLogAlteracao"

Public Sub TABLE_AfterPost()
  RegistrarLogAlteracao "SAM_PRESTADOR_ESPECIALIDADEGRP", CurrentQuery.FieldByName("HANDLE").AsInteger, "TABLE_AfterPost"
End Sub

Public Sub TABLE_AfterScroll()
	If liberaEspecialidade <>"" Then
		ESPECIALIDADE.ReadOnly = True
		ESPECIALIDADEGRUPO.ReadOnly = True
		PERMITEEXECUTAR.ReadOnly = True
		PRESTADOR.ReadOnly = True
		PERMITERECEBER.ReadOnly = True
		PRESTADORESPECIALIDADE.ReadOnly = True
	Else
		ESPECIALIDADE.ReadOnly = False
		ESPECIALIDADEGRUPO.ReadOnly = False
		PERMITEEXECUTAR.ReadOnly = False
		PRESTADOR.ReadOnly = False
		PERMITERECEBER.ReadOnly = False
		PRESTADORESPECIALIDADE.ReadOnly = False
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaEspecialidade

	If Msg <>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If

	Set qPESSOA = NewQuery

	qPESSOA.Add("SELECT * FROM SAM_PRESTADOR WHERE HANDLE=:qPREST")
	qPESSOA.ParamByName("qPREST").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	qPESSOA.Active = True

	Set SQL = NewQuery

	If qPESSOA.FieldByName("FISICAJURIDICA").Value = 1 Then
		SQL.Add("SELECT *											")
		SQL.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID			")
		SQL.Add(" WHERE PRESTADOR = :PRESTADOR 					")
		SQL.Add("   AND DATAINICIAL <= :DATA           			")
		SQL.Add("   AND (DATAFINAL  IS NULL OR DATAFINAL >= :DATA)")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL.ParamByName("DATA").Value = ServerDate
		SQL.Active = True
	Else
		SQL.Add("SELECT *											")
		SQL.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID			")
		SQL.Add(" WHERE ENTIDADE = :ENTIDADE 						")
		SQL.Add("   AND DATAINICIAL <= :DATA           			")
		SQL.Add("   AND (DATAFINAL  IS NULL OR DATAFINAL >= :DATA)")

		SQL.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL.ParamByName("DATA").Value = ServerDate
		SQL.Active = True
	End If

	While Not SQL.EOF
		Dim qESP As Object
		Set qESP = NewQuery

		qESP.Add("SELECT G.HANDLE ")
		qESP.Add("  FROM SAM_MEMBRO_ESPECIALIDADEGRUPO G ")
		qESP.Add("  JOIN SAM_MEMBRO_ESPECIALIDADE      P ON (P.HANDLE = G.MEMBROESPECIALIDADE) ")
		qESP.Add(" WHERE P.CORPOCLINICO = :CORPOCLINICO  ")
		qESP.Add("   AND G.ESPECIALIDADEGRUPO =:ESPECIALIDADEGRUPO ")

		qESP.ParamByName("CORPOCLINICO").Value = SQL.FieldByName("HANDLE").Value
		qESP.ParamByName("ESPECIALIDADEGRUPO").Value = CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").Value
		qESP.Active = False
		qESP.Active = True

		If Not qESP.EOF Then
			CanContinue = False

			Set qPREST = NewQuery

			If qPESSOA.FieldByName("FISICAJURIDICA").Value = 1 Then
				qPREST.Add("SELECT P.NOME ")
				qPREST.Add("  FROM SAM_PRESTADOR P ")
				qPREST.Add("  JOIN SAM_PRESTADOR_PRESTADORDAENTID E ON (E.ENTIDADE = P.HANDLE) ")
				qPREST.Add(" WHERE E.HANDLE = :HANDLE ")

				qPREST.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
				qPREST.Active = False
				qPREST.Active = True

				CanContinue = False

				bsShowmEssage("Este grupo não pode ser excluído !!!" + Chr(10) + _
					"Motivo: Ele está cadastrado na Especialidade do Corpo-Clinico da " + Chr(10) + _
					"entidade " + qPREST.FieldByName("NOME").AsString + "," + Chr(10) + "onde o membro é este prestado.", "E")
			Else
				qPREST.Add("SELECT P.NOME ")
				qPREST.Add("  FROM SAM_PRESTADOR P ")
				qPREST.Add("  JOIN SAM_PRESTADOR_PRESTADORDAENTID E ON (E.PRESTADOR = P.HANDLE) ")
				qPREST.Add(" WHERE E.HANDLE = :HANDLE ")

				qPREST.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
				qPREST.Active = False
				qPREST.Active = True

				CanContinue = False

				bsShowMessage("Este grupo não pode ser excluído !!!" + Chr(10) + "Motivo: Ele está cadastrado no membro do corpo-clinico - " + _
					qPREST.FieldByName("NOME").AsString, "E")
			End If

			Exit Sub
		End If

		SQL.Next
	Wend
	If CanContinue Then
	    RegistrarLogAlteracao "SAM_PRESTADOR_ESPECIALIDADE", CurrentQuery.FieldByName("PRESTADORESPECIALIDADE").AsInteger, "SAM_PRESTADOR_ESPECIALIDADEGRP.TABLE_BeforeDelete"
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaEspecialidade

	If Msg <>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaEspecialidade

	If Msg <>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("PERMITEEXECUTAR").AsString <>"S" And _
	   CurrentQuery.FieldByName("PERMITERECEBER").AsString <>"S" Then
		CanContinue = False
		bsShowMessage("Deve selecionar Permite Executar e/ou Permite Receber", "E")
	End If

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT DATAINICIAL, DATAFINAL FROM SAM_PRESTADOR_ESPECIALIDADE WHERE HANDLE = :PRESTADORESPECIALIDADE")

	SQL.ParamByName("PRESTADORESPECIALIDADE").Value = CurrentQuery.FieldByName("PRESTADORESPECIALIDADE").Value
	SQL.Active = True

	If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAINICIAL").AsDateTime Then
		CanContinue = False
		bsShowMessage("ERRO:  Data inicial não pode ser menor que a Data inicial da Especialidade", "E")
	Else
		If(Not SQL.FieldByName("DATAFINAL").IsNull)And(CurrentQuery.FieldByName("DATAFINAL").IsNull)Then
			CanContinue = False
			bsShowMessage("ERRO:  Data final não pode ser nula, deve ser menor ou igual a Data final da Especialidade", "E")
		Else
			If(Not SQL.FieldByName("DATAFINAL").IsNull)And(CurrentQuery.FieldByName("DATAFINAL").Value >SQL.FieldByName("DATAFINAL").Value)Then
				CanContinue = False
				bsShowMessage("ERRO:  Data final não pode ser maior que a Data final da Especialidade", "E")
			End If
		End If
	End If

	If CanContinue = True Then
		Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

		Condicao = "AND PRESTADORESPECIALIDADE     = " + CurrentQuery.FieldByName("PRESTADORESPECIALIDADE").AsString
		Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ESPECIALIDADEGRP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADEGRUPO", Condicao)

		If Linha = "" Then
			CanContinue = True
		Else
			CanContinue = False
			bsShowMessage(Linha, "E")
		End If

		Set Interface = Nothing
	End If

	Set qPESSOA = NewQuery

	qPESSOA.Add("SELECT * FROM SAM_PRESTADOR WHERE HANDLE=:qPREST")

	qPESSOA.ParamByName("qPREST").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	qPESSOA.Active = True

	Set SQL = NewQuery

	If qPESSOA.FieldByName("FISICAJURIDICA").Value = 1 Then
		SQL.Add("SELECT *											")
		SQL.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID			")
		SQL.Add(" WHERE PRESTADOR = :PRESTADOR 					")
		SQL.Add("   AND DATAINICIAL <= :DATA           			")
		SQL.Add("   AND (DATAFINAL  IS NULL OR DATAFINAL >= :DATA)")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL.ParamByName("DATA").Value = ServerDate
		SQL.Active = True
	Else
		SQL.Add("SELECT *											")
		SQL.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID			")
		SQL.Add(" WHERE ENTIDADE = :ENTIDADE 						")
		SQL.Add("   AND DATAINICIAL <= :DATA           			")
		SQL.Add("   AND (DATAFINAL  IS NULL OR DATAFINAL >= :DATA)")

		SQL.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL.ParamByName("DATA").Value = ServerDate
		SQL.Active = True
	End If

	While Not SQL.EOF
		Dim qESP As Object
		Set qESP = NewQuery

		qESP.Add("SELECT G.HANDLE ")
		qESP.Add("  FROM SAM_MEMBRO_ESPECIALIDADEGRUPO G ")
		qESP.Add("  JOIN SAM_MEMBRO_ESPECIALIDADE      P ON (P.HANDLE = G.MEMBROESPECIALIDADE) ")
		qESP.Add(" WHERE P.CORPOCLINICO = :CORPOCLINICO  ")
		qESP.Add("   AND G.ESPECIALIDADEGRUPO =:ESPECIALIDADEGRUPO ")

		qESP.ParamByName("CORPOCLINICO").Value = SQL.FieldByName("HANDLE").Value
		qESP.ParamByName("ESPECIALIDADEGRUPO").Value = CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").Value
		qESP.Active = False
		qESP.Active = True

		If Not qESP.EOF Then
			CanContinue = False

			Set qPREST = NewQuery

			If qPESSOA.FieldByName("FISICAJURIDICA").Value = 1 Then
				qPREST.Add("SELECT P.NOME ")
				qPREST.Add("  FROM SAM_PRESTADOR P ")
				qPREST.Add("  JOIN SAM_PRESTADOR_PRESTADORDAENTID E ON (E.ENTIDADE = P.HANDLE) ")
				qPREST.Add(" WHERE E.HANDLE = :HANDLE ")

				qPREST.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
				qPREST.Active = False
				qPREST.Active = True

				CanContinue = False

				bsShowMessage("Este grupo não pode ser excluído !!!" + Chr(10) + _
					"Motivo: Ele está cadastrado na Especialidade do Corpo-Clinico da " + Chr(10) + _
					"entidade " + qPREST.FieldByName("NOME").AsString + "," + Chr(10) + "onde o membro é este prestado.", "E")
			Else
				qPREST.Add("SELECT P.NOME ")
				qPREST.Add("  FROM SAM_PRESTADOR P ")
				qPREST.Add("  JOIN SAM_PRESTADOR_PRESTADORDAENTID E ON (E.PRESTADOR = P.HANDLE) ")
				qPREST.Add(" WHERE E.HANDLE = :HANDLE ")

				qPREST.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
				qPREST.Active = False
				qPREST.Active = True

				CanContinue = False

				bsShowMessage("Este grupo não pode ser incluído/excluído !!!" + Chr(10) + _
					"Motivo: Ele está cadastrado no membro do corpo-clinico - " + _
					qPREST.FieldByName("NOME").AsString + Chr(10) + _
					"e esta operação poderá causar inconsistência nos dados do corpo-clínico.", "E")
			End If

		    Exit Sub
		End If

		SQL.Next
	Wend
End Sub
