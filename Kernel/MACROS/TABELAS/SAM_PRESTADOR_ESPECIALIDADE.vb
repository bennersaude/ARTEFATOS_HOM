'HASH: D414CDC08AE6981AE58FB94FC3579FDE
'MACRO: SAM_PRESTADOR_ESPECIALIDADE
'#Uses "*bsShowMessage"
'#Uses "*liberaEspecialidade"
'#Uses "*RegistrarLogAlteracao"
Option Explicit


Dim vESPECIALIDADE As Long
Dim count          As Integer

Public Sub TABLE_AfterCommitted()
	Dim SQL  As Object
	Dim SQL1 As Object
	Dim SQL2 As Object
	Dim SQL3 As Object
	Dim QIns As Object
	Dim SamPrestadorBLL As CSBusinessComponent


	If ((CurrentQuery.FieldByName("PUBLICARNOLIVRO").AsString = "S") And (CurrentQuery.FieldByName("DATAFINAL").IsNull )) Then 'Coelho SMS: 110505
		Set SQL = NewQuery

		SQL.Add("SELECT * FROM SAM_PRESTADOR_LIVRO WHERE PRESTADOR = :PRESTADOR AND ESPECIALIDADE = :ESPECIALIDADE")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
		SQL.Active = True

		If SQL.EOF Then
			Set SQL1 = NewQuery

			SQL1.Clear

			SQL1.Add("SELECT E.HANDLE HENDERECO,              ")
			SQL1.Add("       A.HANDLE HAREA                   ")
			SQL1.Add("  FROM SAM_PRESTADOR_ENDERECO E,        ")
			SQL1.Add("       SAM_DIMENSIONAMENTO    D,        ")
			SQL1.Add("       SAM_AREALIVRO          A         ")
			SQL1.Add(" WHERE A.HANDLE        = D.AREALIVRO    ")
			SQL1.Add("   AND E.PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString)
			SQL1.Add("   AND D.ESPECIALIDADE = " + CurrentQuery.FieldByName("ESPECIALIDADE").AsString)
			SQL1.Add("   AND E.ATENDIMENTO   = 'S'            ")
			SQL1.Add("   AND E.DATACANCELAMENTO  IS NULL      ")

			Set SQL2 = NewQuery

			SQL2.Clear
			SQL2.Add("SELECT COUNT(*) REGISTROS               ")
			SQL2.Add("  FROM SAM_PRESTADOR_ENDERECO E,        ")
			SQL2.Add("       SAM_DIMENSIONAMENTO    D,        ")
			SQL2.Add("       SAM_AREALIVRO          A         ")
			SQL2.Add("  WHERE A.HANDLE        = D.AREALIVRO   ")
			SQL2.Add("   AND E.PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString)
			SQL2.Add("   AND D.ESPECIALIDADE = " + CurrentQuery.FieldByName("ESPECIALIDADE").AsString)
			SQL2.Add("   AND E.ATENDIMENTO   = 'S'            ")
			SQL2.Add("   AND E.DATACANCELAMENTO  IS NULL      ")
			SQL2.Active = True

			If SQL2.FieldByName("REGISTROS").AsInteger >1 Then
		      Dim vsXMLSelecao As String
		      Dim vsMensagem   As String

		      vsXMLSelecao = ""

		      If WebMode Then
		        SQL1.Active = True

		        Dim vcContainer As CSDContainer

		        Set vcContainer = NewContainer

		        vcContainer.GetFieldsFromQuery(SQL1.TQuery)
		        vcContainer.LoadAllFromQuery(SQL1.TQuery)

		        vsXMLSelecao = vcContainer.GetXML

		        Set vcContainer = Nothing
		      Else
		        Dim dllBSInterface0020_SelecaoEnderecoAreaLivro As Object

				Set dllBSInterface0020_SelecaoEnderecoAreaLivro = CreateBennerObject("BSINTERFACE0020.SelecaoEnderecoAreaLivro")

				If dllBSInterface0020_SelecaoEnderecoAreaLivro.Exec(CurrentSystem, _
				                                                    CurrentQuery.FieldByName("PRESTADOR").AsInteger, _
				                                                    CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
				                                                    True, _
				                                                    vsXMLSelecao, _
				                                                    vsMensagem) = 1 Then
                  bsShowMessage("Erro na seleção de endereços e área de livro: " + vsMensagem, "I")
				End If

				Set dllBSInterface0020_SelecaoEnderecoAreaLivro = Nothing
		      End If

			  If vsXMLSelecao <> "" Then
                Dim dllBSPre001_AtualizacaoEspecialidade As Object
                Set dllBSPre001_AtualizacaoEspecialidade = CreateBennerObject("BSPRE001.AtualizacaoEspecialidade")

                If dllBSPre001_AtualizacaoEspecialidade.Livro(CurrentSystem, _
				                                              CurrentQuery.FieldByName("PRESTADOR").AsInteger, _
				                                              CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
				                                              vsXMLSelecao, _
				                                              vsMensagem) Then
                  bsShowMessage("Erro na atualização do livro: " + vsMensagem, "I")
                Else
                  If WebMode Then
                    bsShowMessage("Especialidade incluída no 'Livro de credenciados' para todos os endereços de atendimento ativos do Prestador!", "I")
                  End If
				End If

                Set dllBSPre001_AtualizacaoEspecialidade = Nothing

              End If
			Else
				SQL1.Active = True

				Set SQL3 = NewQuery

				SQL3.Clear

				SQL3.Add("SELECT COUNT(1) aCHOUEND                ")
				SQL3.Add("  FROM SAM_PRESTADOR_ENDERECO E         ")
				SQL3.Add(" WHERE E.PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString)
				SQL3.Add("   AND E.ATENDIMENTO   = 'S'            ")
				SQL3.Add("   AND E.DATACANCELAMENTO  IS NULL      ")

				SQL3.Active = True

				If SQL3.FieldByName("aCHOUEND").AsInteger = 0 Then
					bsShowMessage("Especialidade não adicionada no livro." + Chr(10) + _
								  "Prestador não possui endereço de atendimento ou o endereço de atendimento está cancelado.", "I")
					Exit Sub
				Else
					If SQL1.FieldByName("HAREA").IsNull Then
						bsShowMessage( "Especialidade não adicionada no livro." + Chr(10) + "Especialidade não possui área do livro.","I")
						Exit Sub
					End If
				End If

				Set QIns = NewQuery

				StartTransaction

				QIns.Add("INSERT INTO SAM_PRESTADOR_LIVRO   ")
				QIns.Add("   (HANDLE,                       ")
				QIns.Add("    PRESTADOR,					")
				QIns.Add("    AREA,							")
				QIns.Add("    ESPECIALIDADE,				")
				QIns.Add("    ENDERECO,						")
				QIns.Add("    PUBLICARNOLIVRO,				")
				QIns.Add("    PUBLICARINTERNET,				")
				QIns.Add("    VISUALIZARCENTRAL,			")
				QIns.Add("    OBSERVACAO)					")
				QIns.Add("    VALUES  						")
				QIns.Add("    (:HANDLE,	    				")
				QIns.Add("     :PRESTADOR,					")
				QIns.Add("     :AREA,						")
				QIns.Add("     :ESPECIALIDADE,				")
				QIns.Add("     :ENDERECO,					")
				QIns.Add("     'S',							")
				QIns.Add("     'S',							")
				QIns.Add("     'S',							")
				QIns.Add("     NULL)						")

				QIns.ParamByName("HANDLE").Value = NewHandle("SAM_PRESTADOR_LIVRO")
				QIns.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
				QIns.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
				QIns.ParamByName("AREA").Value = SQL1.FieldByName("HAREA").AsInteger
				QIns.ParamByName("ENDERECO").Value = SQL1.FieldByName("HENDERECO").AsInteger

				QIns.ExecSQL

				Commit

				bsShowMessage("Especialidade incluída no 'Livro de credenciados' !", "I")
			End If
		End If
	End If

	Set SamPrestadorBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.SamPrestadorBLL, Benner.Saude.Prestadores.Business")
	SamPrestadorBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
	SamPrestadorBLL.Execute("VerificarSeExportaBennerHospitalar")

	Set SQL  = Nothing
	Set SQL1 = Nothing
	Set SQL2 = Nothing
	Set SQL3 = Nothing
	Set QIns = Nothing
	Set SamPrestadorBLL = Nothing
End Sub

Public Sub TABLE_AfterPost()
    RegistrarLogAlteracao "SAM_PRESTADOR_ESPECIALIDADE", CurrentQuery.FieldByName("HANDLE").AsInteger, "TABLE_AfterPost"

	Dim SQL As Object

	If CurrentQuery.FieldByName("PRINCIPAL").AsString = "S" Then
		Set SQL = NewQuery

		SQL.Clear

		SQL.Add("UPDATE SAM_PRESTADOR_ESPECIALIDADE SET PRINCIPAL = 'N' WHERE PRESTADOR = :PRESTADOR AND HANDLE <> :HANDLE AND PRINCIPAL = 'S' ")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

		SQL.ExecSQL

		Set SQL = Nothing
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If liberaEspecialidade <>"" Then
		ESPECIALIDADE.ReadOnly = True
		PRESTADOR.ReadOnly = True
		PRINCIPAL.ReadOnly = True
		TEMPORARIO.ReadOnly = True
	Else
		ESPECIALIDADE.ReadOnly = False
		PRESTADOR.ReadOnly = False
		PRINCIPAL.ReadOnly = False
		TEMPORARIO.ReadOnly = False
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If CurrentQuery.FieldByName("PRINCIPAL").Value = "S" Then
	    bsShowMessage("Esta é a especialidade principal. Informar outra antes de remover","E" )
	    CanContinue = False
	    Exit Sub
    End If

	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaEspecialidade

	If Msg <> "" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_PRESTADOR_LIVRO WHERE ESPECIALIDADE = :ESPECIALIDADE ")
	SQL.Add("                                    AND PRESTADOR     = :PRESTADOR     ")

	SQL.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").Value
	SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
	SQL.Active = True

	If Not SQL.EOF Then
		bsShowMessage("Operação Cancelada !!!" + Chr(10) + _
			"Motivo: esta especialidade está cadastrada em 'Livro de Credenciamentos' do prestador", "E")
		CanContinue = False
	End If

	Dim qPESSOA As Object
	Set qPESSOA = NewQuery

	qPESSOA.Add("SELECT * FROM SAM_PRESTADOR WHERE HANDLE=:qPREST")

	qPESSOA.ParamByName("qPREST").Value = RecordHandleOfTable("SAM_PRESTADOR")
	qPESSOA.Active = True

	If qPESSOA.FieldByName("FISICAJURIDICA").Value = 1 Then
		SQL.Clear

		SQL.Add("SELECT *												")
		SQL.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID				")
		SQL.Add(" WHERE PRESTADOR = :PRESTADOR 						")
		SQL.Add("   AND DATAINICIAL <= :DATA           				")
		SQL.Add("   AND (DATAFINAL  IS NULL OR DATAFINAL >= :DATA)	")

		SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
		SQL.ParamByName("DATA").Value = ServerDate
		SQL.Active = True
	Else
		SQL.Clear

		SQL.Add("SELECT *											")
		SQL.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID				")
		SQL.Add(" WHERE ENTIDADE = :ENTIDADE 						")
		SQL.Add("   AND DATAINICIAL <= :DATA           				")
		SQL.Add("   AND (DATAFINAL  IS NULL OR DATAFINAL >= :DATA)	")

		SQL.ParamByName("ENTIDADE").Value = RecordHandleOfTable("SAM_PRESTADOR")
		SQL.ParamByName("DATA").Value = ServerDate
		SQL.Active = True
	End If

	Dim qESP As Object
	Set qESP = NewQuery
	Dim qPREST As Object
	Set qPREST = NewQuery

	While Not SQL.EOF
		qESP.Clear
		qESP.Add("SELECT * FROM SAM_MEMBRO_ESPECIALIDADE WHERE CORPOCLINICO = :CORPOCLINICO AND ESPECIALIDADE =:ESPECIALIDADE")
		qESP.ParamByName("CORPOCLINICO").Value = SQL.FieldByName("HANDLE").Value
		qESP.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").Value
		qESP.Active = True

		If Not qESP.EOF Then
			CanContinue = False
			qPREST.Clear
			If qPESSOA.FieldByName("FISICAJURIDICA").Value = 1 Then
				qPREST.Add("SELECT P.NOME ")
				qPREST.Add("  FROM SAM_PRESTADOR P ")
				qPREST.Add("  JOIN SAM_PRESTADOR_PRESTADORDAENTID E ON (E.ENTIDADE = P.HANDLE) ")
				qPREST.Add(" WHERE E.HANDLE = :HANDLE ")

				qPREST.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
				qPREST.Active = False
				qPREST.Active = True
				bsShowMessage("Esta especialidade não pode ser excluída !!!" + Chr(10) + _
					"Motivo: Ela está cadastrada na Especialidade do Corpo-Clinico da " + Chr(10) + _
					"entidade " + qPREST.FieldByName("NOME").AsString + "," + Chr(10) + "onde o membro é este prestado.", "E")
			Else
				qPREST.Add("SELECT P.NOME ")
				qPREST.Add("  FROM SAM_PRESTADOR P ")
				qPREST.Add("  JOIN SAM_PRESTADOR_PRESTADORDAENTID E ON (E.PRESTADOR = P.HANDLE) ")
				qPREST.Add(" WHERE E.HANDLE = :HANDLE ")

				qPREST.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
				qPREST.Active = False
				qPREST.Active = True

				bsShowMessage("Esta especialidade não pode ser excluída !!!" + Chr(10) + _
					"Motivo: Ela está cadastrada no membro do corpo-clinico - " + _
					qPREST.FieldByName("NOME").AsString, "E")
			End If
			Set qPESSOA = Nothing
			Set qESP    = Nothing
			Set SQL     = Nothing
			Set qPREST  = Nothing
			Exit Sub
		End If

		SQL.Next
	Wend

	If CanContinue Then
        RegistrarLogAlteracao "SAM_PRESTADOR", CurrentQuery.FieldByName("PRESTADOR").AsInteger, "SAM_PRESTADOR_ESPECIALIDADE.TABLE_BeforeDelete"
	End If

	Set qPESSOA = Nothing
	Set qESP    = Nothing
	Set SQL     = Nothing
	Set qPREST  = Nothing
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

	If Not CurrentQuery.FieldByName("ESPECIALIDADE").IsNull Then
		vESPECIALIDADE = CurrentQuery.FieldByName("ESPECIALIDADE").Value
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String
	Msg = LiberaEspecialidade

	If Msg <>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Msg As String

	count = 0

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
	'SMS 49152 - Fim

	Dim qGRUPO As Object
	Dim qREDE As Object
	Dim qSUB As Object
	Dim qLIVRO As Object
	Dim qAtualiza As Object
	Dim qMEMBRO As Object
	Dim qAux As Object

	Set qAtualiza = NewQuery

	Dim vAbertas As Integer 'coelho
    vAbertas = 0

	If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <= CDate("01/01/1900")Then
		CanContinue = False
		bsShowMessage("Data inicial inválida !!!", "E")
		Exit Sub
	End If

	'------se a especialidade está sendo alterada verificar se as restrições --------
	'------ou seja verificar se esta foi relacionada em alguma tabela do prestador --
	Set qGRUPO = NewQuery
	Set qREDE = NewQuery
	Set qSUB = NewQuery
	Set qLIVRO = NewQuery
	Set qMEMBRO = NewQuery
	Set qAux = NewQuery

	If vESPECIALIDADE <>CurrentQuery.FieldByName("ESPECIALIDADE").Value Then
		'---grupo ---
		qGRUPO.Add("SELECT * FROM SAM_PRESTADOR_ESPECIALIDADEGRP WHERE PRESTADORESPECIALIDADE = :PRESTADORESPECIALIDADE")
		qGRUPO.Add(" AND DATAINICIAL <= :DATA AND (DATAFINAL IS NULL OR DATAFINAL >= :DATA)                            ")

		qGRUPO.ParamByName("PRESTADORESPECIALIDADE").Value = CurrentQuery.FieldByName("HANDLE").Value
		qGRUPO.ParamByName("DATA").Value = ServerDate
		qGRUPO.Active = True

		If Not qGRUPO.EOF Then
			CanContinue = False
			bsShowMessage("O Campo Especialidade não pode ser alterado" + Chr(10) + _
				"Motivo: Existem registros na carga 'Grupo de Eventos da Especialidade'", "E")
			Exit Sub
		End If

		'---especialidade ---
		qREDE.Add("SELECT * FROM SAM_PRESTADOR_ESPEC_REDE WHERE PRESTADORESPECIALIDADE = :PRESTADORESPECIALIDADE")

		qREDE.ParamByName("PRESTADORESPECIALIDADE").Value = CurrentQuery.FieldByName("HANDLE").Value
		qREDE.Active = True

		If Not qREDE.EOF Then
			CanContinue = False
			bsShowMessage("O Campo Especialidade não pode ser alterado" + Chr(10) + _
				"Motivo: Existem registros na carga 'Rede Restrita da Especialidade'", "E")
			Exit Sub
		End If

		'--na subespecialidade ---
		qSUB.Add("SELECT * FROM SAM_PRESTADOR_ESPECIALIDADESUB WHERE PRESTADORESPECIALIDADE = :PRESTADORESPECIALIDADE")

		qSUB.ParamByName("PRESTADORESPECIALIDADE").Value = CurrentQuery.FieldByName("HANDLE").Value
		qSUB.Active = True

		If Not qSUB.EOF Then
			CanContinue = False
			bsShowMessage("O Campo Especialidade não pode ser alterado" + Chr(10) + _
				"Motivo: Existem registros na carga 'Sub-Especialidade'", "E")
			Exit Sub
		End If

		'---no livro de credenciados ---
		qLIVRO.Add("SELECT * FROM SAM_PRESTADOR_LIVRO WHERE ESPECIALIDADE = :ESPECIALIDADE ")
		qLIVRO.Add("                                    AND PRESTADOR     = :PRESTADOR     ")

		qLIVRO.ParamByName("ESPECIALIDADE").Value = vESPECIALIDADE
		qLIVRO.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
		qLIVRO.Active = True

		If Not qLIVRO.EOF Then
			CanContinue = False
			bsShowMessage("O Campo Especialidade não pode ser alterado" + Chr(10) + _
				"Motivo: Esta especialidade está cadastrada em 'Livro de Credenciamentos' do prestador", "E")
			Exit Sub
		End If
	End If

	If vESPECIALIDADE <>CurrentQuery.FieldByName("ESPECIALIDADE").Value Or Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		'---nos membros do corpo-clinico ---
		qAux.Add("SELECT * FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")

		qAux.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		qAux.Active = True

		If qAux.FieldByName("FISICAJURIDICA").Value = 1 Then
			qMEMBRO.Clear

			qMEMBRO.Add("SELECT *																	")
			qMEMBRO.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID P		                    		")
			qMEMBRO.Add("  JOIN SAM_MEMBRO_ESPECIALIDADE       M ON (M.CORPOCLINICO = P.HANDLE)   ")
			qMEMBRO.Add(" WHERE P.PRESTADOR     = :PRESTADOR 										")
			qMEMBRO.Add("   AND M.ESPECIALIDADE = :ESPECIALIDADE 									")
			qMEMBRO.Add("   AND P.DATAINICIAL <= :DATA           									")
			qMEMBRO.Add("   AND (P.DATAFINAL  IS NULL OR P.DATAFINAL >= :DATA)					")

			qMEMBRO.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
			qMEMBRO.ParamByName("ESPECIALIDADE").Value = vESPECIALIDADE
			qMEMBRO.ParamByName("DATA").Value = ServerDate
			qMEMBRO.Active = False
			qMEMBRO.Active = True

			If Not qMEMBRO.EOF Then
				qAux.Clear

				qAux.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE ")
				qAux.ParamByName("HANDLE").Value = qMEMBRO.FieldByName("ENTIDADE").Value
				qAux.Active = True

				CanContinue = False

				bsShowMessage("O Campo Especialidade não pode ser alterado ou ter a vigência fechada !" + Chr(10) + _
					"Motivo: Esta Especialidade está cadastrada no Corpo-Clinico da " + Chr(10) + _
					"entidade " + qAux.FieldByName("NOME").AsString + "," + Chr(10) + _
					"onde o membro é este prestador.", "E")
				Exit Sub
			End If
		Else
			qMEMBRO.Clear

			qMEMBRO.Add("SELECT *																	")
			qMEMBRO.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID P									")
			qMEMBRO.Add("  JOIN SAM_MEMBRO_ESPECIALIDADE       M ON (M.CORPOCLINICO = P.HANDLE)   ")
			qMEMBRO.Add(" WHERE P.ENTIDADE     = :ENTIDADE 										")
			qMEMBRO.Add("   AND M.ESPECIALIDADE = :ESPECIALIDADE  								")
			qMEMBRO.Add("   AND P.DATAINICIAL <= :DATA           									")
			qMEMBRO.Add("   AND (P.DATAFINAL  IS NULL OR P.DATAFINAL >= :DATA)					")

			qMEMBRO.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
			qMEMBRO.ParamByName("ESPECIALIDADE").Value = vESPECIALIDADE
			qMEMBRO.ParamByName("DATA").Value = ServerDate
			qMEMBRO.Active = False
			qMEMBRO.Active = True
		End If

		If Not qMEMBRO.EOF Then
			qAux.Clear

			qAux.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE ")

			qAux.ParamByName("HANDLE").Value = qMEMBRO.FieldByName("PRESTADOR").Value
			qAux.Active = True

			CanContinue = False

			bsShowMessage("O Campo Especialidade não pode ser alterado ou ter a vigência fechada !" + Chr(10) + _
				"Motivo: Ela está cadastrada no membro do corpo-clinico - " + qAux.FieldByName("NOME").AsString, "E")
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		If CurrentQuery.FieldByName("DATAINICIAL").Value >CurrentQuery.FieldByName("DATAFINAL").Value Then
			bsShowMessage("A Data Inicial não pode ser maior que a Data Final", "E")
			CanContinue = False
		End If
	End If

	Dim Interface As Object
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Dim Condicao As String
	Dim Linha As String

	Condicao = "AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ESPECIALIDADE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADE", Condicao)

	If Linha <>"" Then
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing

	Dim SQL As Object

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		Set SQL = NewQuery

		SQL.Add("SELECT DATAINICIAL, DATAFINAL FROM SAM_PRESTADOR_ESPECIALIDADEGRP WHERE PRESTADORESPECIALIDADE = :HANDLE")

		SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
		SQL.Active = True

		While Not SQL.EOF
			If SQL.FieldByName("DATAFINAL").IsNull Then
				'bsShowMessage("A data final não pode ser preenchida - existe(m) grupo(s) com vigência(s) aberta(s)", "E")
				'CanContinue = False
				vAbertas = vAbertas + 1 ' Coelho SMS: 110505
			Else
				If SQL.FieldByName("DATAFINAL").Value >CurrentQuery.FieldByName("DATAFINAL").Value Then
					bsShowMessage("A data final não pode ser preenchida - existe(m) grupo(s) com data final maior", "E")
					CanContinue = False
				Else
					If SQL.FieldByName("DATAINICIAL").Value >CurrentQuery.FieldByName("DATAFINAL").Value Then
						bsShowMessage("A data final não pode ser preenchida - existe(m) grupo(s) com data inicial maior", "E")
						CanContinue = False
					End If
				End If
			End If

			SQL.Next
		Wend

	End If

	Set SQL = Nothing

	If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And (CanContinue = True) Then
		Dim MEMBRO As Object
		Set SQL = NewQuery

		SQL.Add("SELECT HANDLE FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE ENTIDADE = :PRESTADOR")
		SQL.Add("AND DATAINICIAL >= :DATA AND (DATAFINAL IS NULL OR DATAFINAL <= :DATA)       ")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
		SQL.ParamByName("DATA").Value = ServerDate
		SQL.Active = True

		While Not SQL.EOF
			Set MEMBRO = NewQuery

			MEMBRO.Add("SELECT HANDLE FROM SAM_MEMBRO_ESPECIALIDADE WHERE CORPOCLINICO = :CORPOCLINICO AND ESPECIALIDADE =:ESPECIALIDADE")

			MEMBRO.ParamByName("CORPOCLINICO").Value = SQL.FieldByName("HANDLE").Value
			MEMBRO.ParamByName("ESPECIALIDADE").Value = vESPECIALIDADE
			MEMBRO.Active = True

			If Not MEMBRO.FieldByName("HANDLE").IsNull Then
				CanContinue = False

				Msg = "Vigência não pode ser fechada !!! " + Chr(13)
				Msg = Msg + "Esta especialidade está cadastrada em membros do corpo-clínico."

				bsShowMEssage(Msg, "E")
			End If

			Set MEMBRO = Nothing

			SQL.Next
		Wend

		Set SQL = Nothing
	End If

	'Matiello -20/07/2007 ---------------------------------------------------------------------
	'SMS 84849 - Verifica se o prestador tem alguma especialidade principal caso seu tipo exija.
	Dim qrPessoa As Object 'tipo do prestador
	Dim qrTemEsp As Object 'verifica especialidade

	'busca tipo do prestador
	Set qrPessoa =NewQuery

	qrPessoa.Add("SELECT NOME, TIPOPRESTADOR " + _
				 "  FROM SAM_PRESTADOR       " + _
				 " WHERE HANDLE = " + CurrentQuery.FieldByName("PRESTADOR").AsString)

	qrPessoa.Active =True

	If CurrentQuery.FieldByName("PRINCIPAL").Value <>"S" Then
		'exige o tipo do prestador
		If qrPessoa.FieldByName("TIPOPRESTADOR").IsNull Then
			bsShowMessage("Cadastre o tipo para este prestador" ,"E")
			CanContinue =False
			Set qrPessoa =Nothing
			Exit Sub
		End If
		Set qrPessoa =Nothing

		'busca ao menos uma especialidade principal
		Set qrTemEsp =NewQuery

		qrTemEsp.Add("SELECT HANDLE, PRESTADOR, ESPECIALIDADE, PRINCIPAL ")
		qrTemEsp.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE   	    ")
		qrTemEsp.Add("  WHERE (DATAFINAL Is Null Or DATAFINAL >= :DATA)")
		qrTemEsp.Add("   AND PRINCIPAL = 'S' ")
		qrTemEsp.Add("   AND PRESTADOR = :HPRESTADOR")
		qrTemEsp.Add("   AND HANDLE    <>:HPRESTADORESPECIALIDADE")
        qrTemEsp.ParamByName("HPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
        qrTemEsp.ParamByName("HPRESTADORESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qrTemEsp.ParamByName("DATA").AsDateTime = ServerDate

		qrTemEsp.Active =True

		If qrTemEsp.EOF Then
  		  bsShowMessage("Operação cancelada!!!  Prestador deve ter ao menos uma especialidade principal" ,"E")
  		  CanContinue = False
		  Set qrTemEsp = Nothing
 		  Exit Sub
 		Else
  		  If vAbertas > 0 Then ' Coelho SMS: 110505
            If bsShowMessage("Existe(m) grupo(s) com vigência(s) aberta(s), deseja encerrá-la(s)?", "Q") = vbYes Then
              qAux.Clear
              qAux.Add("SELECT HANDLE FROM SAM_PRESTADOR_ESPECIALIDADEGRP WHERE PRESTADORESPECIALIDADE = :PPRESTADORESPECIALIDADE AND DATAFINAL IS NULL")
              qAux.ParamByName("PPRESTADORESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
              qAux.Active = True

              qAtualiza.Clear
              qAtualiza.Add("UPDATE SAM_PRESTADOR_ESPECIALIDADEGRP SET DATAFINAL=:PDATAFINAL WHERE HANDLE = :PHANDLE AND DATAFINAL IS NULL")

              While Not qAux.EOF
                qAtualiza.ParamByName("PHANDLE").AsInteger = qAux.FieldByName("HANDLE").AsInteger
                qAtualiza.ParamByName("PDATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
                qAtualiza.ExecSQL

                qAux.Next
              Wend
            Else
              bsShowMessage("A data final não pode ser preenchida - existe(m) grupo(s) com vigência(s) aberta(s)", "E")
              CanContinue = False
            End If
          End If
		End If

		Set qrTemEsp =Nothing
	Else
		Dim qVerificaPrincipal  As Object
		Dim vbAlterarPrincipal  As Boolean
		Dim vsPrincipalAnterior As String
		Set qVerificaPrincipal = NewQuery

        qVerificaPrincipal.Clear
		qVerificaPrincipal.Add("SELECT E.DESCRICAO                                        ")
		qVerificaPrincipal.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE PE  	                  ")
		qVerificaPrincipal.Add("  JOIN SAM_ESPECIALIDADE E ON E.HANDLE = PE.ESPECIALIDADE ")
		qVerificaPrincipal.Add("  WHERE (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= :DATA)    ")
		qVerificaPrincipal.Add("   AND PE.PRINCIPAL = 'S'                                 ")
		qVerificaPrincipal.Add("   AND PE.PRESTADOR = :PRESTADOR                          ")
		If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
		  qVerificaPrincipal.Add("   AND PE.HANDLE <> :HANDLEATUAL                           ")
		  qVerificaPrincipal.ParamByName("HANDLEATUAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		End If
		qVerificaPrincipal.ParamByName("DATA").AsDateTime = ServerDate
		qVerificaPrincipal.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  	    qVerificaPrincipal.Active = True

		If Not qVerificaPrincipal.EOF Then ' se já existe uma marcada como principal não deixa colocar outra
		  vsPrincipalAnterior = qVerificaPrincipal.FieldByName("DESCRICAO").AsString
          If WebMode Then
            vbAlterarPrincipal = True
          Else
            If bsShowMessage("Alterar essa especialidade como principal?","Q" ) = vbYes Then
              vbAlterarPrincipal = True
            Else
              vbAlterarPrincipal = False
            End If
          End If

   		  If vbAlterarPrincipal Then
  	  	    qVerificaPrincipal.Clear
			qVerificaPrincipal.Add("UPDATE SAM_PRESTADOR_ESPECIALIDADE              ")
			qVerificaPrincipal.Add("   SET PRINCIPAL = 'N'           	            ")
			qVerificaPrincipal.Add(" WHERE DATAINICIAL <= :DATA                     ")
			qVerificaPrincipal.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :DATA)")
			qVerificaPrincipal.Add("   AND PRINCIPAL = 'S'                          ")
			qVerificaPrincipal.Add("   AND PRESTADOR = :PRESTADOR                   ")
			If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
		  	  qVerificaPrincipal.Add("   AND HANDLE <> :HANDLEATUAL                 ")
		      qVerificaPrincipal.ParamByName("HANDLEATUAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		    End If
		    qVerificaPrincipal.ParamByName("DATA").AsDateTime = ServerDate
		    qVerificaPrincipal.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
   		    qVerificaPrincipal.ExecSQL

   		    If WebMode Then
              bsShowMessage("Especialidade atual substituiu a especialidade " + vsPrincipalAnterior + " como ""Principal""", "I")
   		    End If

 	        If vAbertas > 0 Then ' Coelho SMS: 110505
              If bsShowMessage("Existe(m) grupo(s) com vigência(s) aberta(s), deseja encerrá-la(s)?", "Q") = vbYes Then
              qAux.Clear
              qAux.Add("SELECT HANDLE FROM SAM_PRESTADOR_ESPECIALIDADEGRP WHERE PRESTADORESPECIALIDADE = :PPRESTADORESPECIALIDADE AND DATAFINAL IS NULL")
              qAux.ParamByName("PPRESTADORESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
              qAux.Active = True

              qAtualiza.Clear
              qAtualiza.Add("UPDATE SAM_PRESTADOR_ESPECIALIDADEGRP SET DATAFINAL=:PDATAFINAL WHERE HANDLE = :PHANDLE AND DATAFINAL IS NULL")

              While Not qAux.EOF
                qAtualiza.ParamByName("PHANDLE").AsInteger = qAux.FieldByName("HANDLE").AsInteger
                qAtualiza.ParamByName("PDATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
                qAtualiza.ExecSQL

                qAux.Next
              Wend
            Else
              bsShowMessage("A data final não pode ser preenchida - existe(m) grupo(s) com vigência(s) aberta(s)", "E")
              CanContinue = False
            End If
            End If
          Else
            CanContinue = False
            CurrentQuery.FieldByName("PRINCIPAL").AsString = "N"
            Exit Sub
          End If
        Else
          ' Coelho SMS: 110505 incluido o else, se não existir outra especialidade como principal, então aborta o processo
          If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
            bsShowMEssage("Processo cancelado! Não existe outra especialidade cadastrada como principal! Se for a primeira, deve ser sem data final!", "E")
            CanContinue = False
            Exit Sub
          End If
		End If

		qVerificaPrincipal.Active = False

		Set qVerificaPrincipal = Nothing
	End If

	VerificarEdital(CanContinue)

	'-------------------------------------------------------------
	'FIM SMS 84849

	'---Claudemir 29.09.2003 -sms 18910
	qLIVRO.Active = False

	qLIVRO.Clear

	qLIVRO.Add("SELECT * FROM SAM_PRESTADOR_LIVRO WHERE ESPECIALIDADE = :ESPECIALIDADE ")
	qLIVRO.Add("                                    AND PRESTADOR     = :PRESTADOR     ")

	qLIVRO.ParamByName("ESPECIALIDADE").Value = vESPECIALIDADE
	qLIVRO.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
	qLIVRO.Active = True

	If Not qLIVRO.EOF Then

        If CurrentQuery.FieldByName("PUBLICARNOLIVRO").AsString = "N" And qLIVRO.FieldByName("PUBLICARNOLIVRO").AsString = "S" Then  ' Paulo Melo - SMS 115572 - 03/06/2009
          bsShowMessage("Campo 'Publicar no livro' não pode ser desmarcado !" + Chr(10) + _
            "Motivo: Esta especialidade está cadastrada em 'Livro de Credenciamentos' do prestador.", "E")
          CurrentQuery.FieldByName("PUBLICARNOLIVRO").Value = "S"
        End If

		If CurrentQuery.FieldByName("PUBLICARINTERNET").AsString = "N" And qLIVRO.FieldByName("PUBLICARINTERNET").AsString = "S" Then
			bsShowMEssage("Campo 'Publicar no internet' não pode ser desmarcado !" + Chr(10) + _
				"Motivo: Esta especialidade está cadastrada em 'Livro de Credenciamentos' do prestador com este campo marcado.", "E")

			CurrentQuery.FieldByName("PUBLICARINTERNET").Value = "S"
		End If

		If CurrentQuery.FieldByName("VISUALIZARCENTRAL").AsString = "N" And qLIVRO.FieldByName("VISUALIZARCENTRAL").AsString = "S" Then
			bsShowMessage("Campo 'Visualizar na central de atendimento' não pode ser desmarcado !" + Chr(10) + _
				"Motivo: Esta especialidade está cadastrada em 'Livro de Credenciamentos' do prestador com este campo marcado.", "E")

			CurrentQuery.FieldByName("VISUALIZARCENTRAL").Value = "S"
		End If
	End If

	Set qGRUPO = Nothing
	Set qREDE = Nothing
	Set qSUB = Nothing
	Set qLIVRO = Nothing
	Set qAtualiza = Nothing
	Set qMEMBRO = Nothing
	Set qAux = Nothing
End Sub

Public Sub VerificarEdital(CanContinue As Boolean)
	If (CurrentQuery.State = 3) Then
	 	If (CanContinue = True) Then
		    Dim componente As CSBusinessComponent

		    On Error GoTo Fim
		  		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.Especialidade.SamEspecialidadeBLL, Benner.Saude.Prestadores.Business")
				componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger)
		  		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
		  		componente.Execute("VerificarCredenciamentoComEdital")
		  	Fim:
				If Err.Number <> 0 Then
	      			Err.Raise(vbsUserException, Err.Source,  "" + Err.Description )
	      			CanContinue = False
	    		End If

	  			Set componente = Nothing
	  	End If
	End If
End Sub
