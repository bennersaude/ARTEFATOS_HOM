'HASH: EE738B15854F12CE77DC45A340A0FFE7

'MACRO TABELA: SAM_PRESTADOR_PROC_MEM_ESP
Option Explicit
'#Uses "*bsShowMessage"

Dim vCondicao As String

Public Sub Condicao()

    CurrentQuery.UpdateRecord

	Dim qMEM As Object
	Set qMEM = NewQuery

	Dim SQL As Object
	Set SQL = NewQuery

	qMEM.Clear
	qMEM.Add("SELECT MEMBRO, PRESTADOR              ")
	qMEM.Add("  FROM SAM_PRESTADOR_PROC_MEMBROS     ")
	qMEM.Add(" WHERE HANDLE = :HANDLE               ")
	qMEM.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_MEMBROS")

	qMEM.Active = True



    Call RetornaEspecialidade(SQL,qMEM.FieldByName("PRESTADOR").AsInteger, qMEM.FieldByName("MEMBRO").AsInteger,  CurrentQuery.FieldByName("TABTIPOMOVIMENTACAO").AsInteger)

	If VisibleMode Then
		vCondicao = "SAM_ESPECIALIDADE.HANDLE"
	End If

	vCondicao = vCondicao + " IN (SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE"
	vCondicao = vCondicao + "      WHERE ESPECIALIDADE     = " + SQL.FieldByName("ESPECIALIDADE").AsInteger

	SQL.Next

	While Not SQL.EOF
		vCondicao = vCondicao + "       OR ESPECIALIDADE      = " + SQL.FieldByName("ESPECIALIDADE").AsInteger
		SQL.Next
	Wend

	vCondicao = vCondicao + ")"


	If VisibleMode Then
		ESPECIALIDADE.LocalWhere = vCondicao
		ESPECIALIDADEINCLUSAO.LocalWhere = vCondicao
	End If

	Set SQL = Nothing
	Set qMEM = Nothing
End Sub

Public Sub TABLE_AfterEdit()
	Condicao
End Sub

Public Sub TABLE_AfterPost()
	Dim componente As CSBusinessComponent

	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcMemEspBLL, Benner.Saude.Prestadores.Business")

	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.Execute("AtualizarEspecialidadesMembro")

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	VerificaEdicao(CanContinue)
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
   VerificaEdicao(CanContinue)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	VerificaEdicao(CanContinue)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("TABTIPOMOVIMENTACAO").AsInteger = 1 Then

		If (CurrentQuery.FieldByName("ESPECIALIDADEINCLUSAO").AsInteger = 0) Then
			bsShowmessage("Campo 'Especialidade' é obrigatório!", "I")
			CanContinue = False
		Else
			CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADEINCLUSAO").AsInteger
		End If

    End If

	VerificarDuplicidadeEspecialidade(CanContinue)
	VerificarMembroCredenciado(CanContinue)
	VerificarEspecialidadeComVigenciaFutura(CanContinue)

End Sub

Public Sub RetornaEspecialidade (ByVal SQL As BPesquisa, Prestador As Integer, Membro As Integer, Tipo As Integer)
		SQL.Clear
		SQL.Add("SELECT * ")
		SQL.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE A")
		SQL.Add(" WHERE A.PRESTADOR = :ENTIDADE")
		SQL.Add("   AND A.DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")

		If Tipo = 1 Then
			SQL.Add("   AND NOT EXISTS (SELECT * ")
		Else
			SQL.Add("   AND EXISTS (SELECT * ")
		End If

		SQL.Add("                 FROM SAM_PRESTADOR_ESPECIALIDADE  B")
		SQL.Add("                WHERE B.DATAINICIAL <= :DATA AND (B.DATAFINAL >= :DATA OR B.DATAFINAL IS NULL)")
		SQL.Add("                  AND B.ESPECIALIDADE = A.ESPECIALIDADE")
		SQL.Add("                  AND B.PRESTADOR = :PRESTADOR)")
		SQL.ParamByName("ENTIDADE").Value = Prestador
		SQL.ParamByName("PRESTADOR").Value = Membro
		SQL.ParamByName("DATA").Value = ServerDate

		SQL.Active = True
End Sub

Public Sub VerificaEdicao(CanContinue As Boolean)
   Dim vMensagem As String
   Dim SQL As Object
   Set SQL = NewQuery

   SQL.Add("SELECT SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.FILIALPADRAO FROM SAM_PRESTADOR_PROC, SAM_PRESTADOR WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PROC.PRESTADOR")
   SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
   SQL.Active = True

   If Not SQL.FieldByName("DATAFINAL").IsNull Then
   	 vMensagem = "Processo finalizado!" + Chr(13)
   End If
   If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
     vMensagem = vMensagem + "Usuário não é o responsável!"
   End If
   Set SQL = Nothing

   If vMensagem <> "" Then
     bsShowMessage(vMensagem, "E")
     CanContinue = False
   End If
End Sub

Public Sub VerificarDuplicidadeEspecialidade(CanContinue As Boolean)
	Dim componente As CSBusinessComponent

	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcMemEspBLL, Benner.Saude.Prestadores.Business")

	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADORPROCMEMBRO").AsInteger)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger)
	componente.AddParameter(pdtAutomatic, CurrentQuery.State = 3)

	If componente.Execute("VerificarDuplicidadeEspecialidade") Then
		bsShowMessage("Especialidade '"& ESPECIALIDADE.Text &"' já cadastrada neste processo!", "E")
		CanContinue = False
	End If

End Sub

Public Sub VerificarMembroCredenciado(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("TABTIPOMOVIMENTACAO").AsInteger = 1) Then

		Dim componente As CSBusinessComponent

		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcMemEspBLL, Benner.Saude.Prestadores.Business")

		componente.AddParameter(pdtInteger, ConsultarPrestadorMembro)

		If componente.Execute("VerificarMembroCredenciado") Then
			bsShowMessage("Não é permitido incluir especialidade para membro do corpo clínico credenciado!", "E")
			CanContinue = False
		End If

	End If
End Sub

Public Sub VerificarEspecialidadeComVigenciaFutura(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("TABTIPOMOVIMENTACAO").AsInteger = 1) Then

		Dim componente As CSBusinessComponent
		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.Especialidade.SamPrestadorEspecialidadeBLL, Benner.Saude.Prestadores.Business")

		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("ESPECIALIDADEINCLUSAO").AsInteger)
		componente.AddParameter(pdtInteger, ConsultarPrestadorMembro)

		If componente.Execute("ExisteEspecialidadeComVigenciaFutura") Then
			bsShowMessage("Não é possível incluir a especialdiade '"+ ESPECIALIDADEINCLUSAO.Text +"'. O Corpo Clínico já possui a especialidade incluída com vigencia futura!", "E")
			CanContinue = False
		End If
	End If
End Sub

Public Function ConsultarPrestadorMembro As Long
	Dim vMembro As Long

	Dim qBusca As Object
	Set qBusca = NewQuery

	qBusca.Add("SELECT MEMBRO                     ")
	qBusca.Add("  FROM SAM_PRESTADOR_PROC_MEMBROS ")
	qBusca.Add(" WHERE HANDLE = :PROCMEMBRO       ")

	qBusca.ParamByName("PROCMEMBRO").AsInteger = CurrentQuery.FieldByName("PRESTADORPROCMEMBRO").AsInteger
	qBusca.Active = True

	vMembro = qBusca.FieldByName("MEMBRO").AsInteger
	Set qBusca = Nothing
	ConsultarPrestadorMembro = vMembro
End Function

Public Sub TABLE_NewRecord()
	CurrentQuery.FieldByName("RESPONSAVELANALISE").Value = CurrentUser
End Sub

Public Sub TABTIPOMOVIMENTACAO_OnChange()
	Condicao
End Sub
