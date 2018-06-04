'HASH: 3612901BC387F67560DE815578BD5FC7
'#uses "*Biblioteca"

Option Explicit

Dim qHandleRelatorio As BPesquisa

Public Sub Main

	Set qHandleRelatorio = NewQuery

	qHandleRelatorio.Add("SELECT HANDLE, NOME FROM R_RELATORIOS WHERE CODIGO =:CODIGO")
	qHandleRelatorio.ParamByName("CODIGO").AsString = SessionVar("CODIGO")
	qHandleRelatorio.Active = True

	If Not qHandleRelatorio.EOF Then

		Select Case SessionVar("CODIGO")

			Case "CAIXA011"

				Call Processo_Emitir_Relatorio(qHandleRelatorio.FieldByName("HANDLE").AsInteger, "G.DATAATENDIMENTO BETWEEN '" & ServerContainer.Field("DATAINICIAL").AsDateTime & "' AND '" & ServerContainer.Field("DATAFINAL").AsDateTime & "'" & _
				IIf(ServerContainer.Field("ESTADO").AsString<>""," AND " &  FilterFieldResultSQL("E.HANDLE",ServerContainer.Field("ESTADO").AsString) ,"") & _
				IIf(ServerContainer.Field("FILIAL").AsString<>""," AND " &  FilterFieldResultSQL("F.HANDLE",ServerContainer.Field("FILIAL").AsString) ,"") & _
				IIf(ServerContainer.Field("PRESTADOR").AsString<>""," AND " &  FilterFieldResultSQL("P.HANDLE",ServerContainer.Field("PRESTADOR").AsString) ,"") & _
				IIf(ServerContainer.Field("LOCALATENDIMENTO").AsString<>""," AND " &  FilterFieldResultSQL("LA.HANDLE",ServerContainer.Field("LOCALATENDIMENTO").AsString) ,"") & _
				" ")

			Case Else
				Call Processo_Emitir_Relatorio(qHandleRelatorio.FieldByName("HANDLE").AsInteger, SessionVar("DefaultWhere"))

		End Select

	End If

	Set qHandleRelatorio = Nothing

End Sub


Public Sub Processo_Emitir_Relatorio(Handle_Relatorio As Long, String_Filtro As String )

		Dim relatorio As CSReportPrinter
		Set relatorio = NewReport(Handle_Relatorio)




	'filtro do relatório---------------------------------------------------------------------

		relatorio.SqlWhere = String_Filtro

    '----------------------------------------------------------------------------------------

		Dim qInsere As Object
		Set qInsere = NewQuery

		qInsere.Add("INSERT INTO R_RELATORIOS_GERADOS (HANDLE, MODULO, USUARIO, SITUACAO, INICIOPROCESSO)")
		qInsere.Add("VALUES")
		qInsere.Add("                                   (:HANDLE, :MODULO, :USUARIO, :SITUACAO, :INICIOPROCESSO)")


		Dim vHandleInsercao As Long

		vHandleInsercao = NewHandle("R_RELATORIOS_GERADOS")

		qInsere.ParamByName("HANDLE").AsInteger = vHandleInsercao
		qInsere.ParamByName("USUARIO").AsInteger = CurrentUser
		qInsere.ParamByName("SITUACAO").AsInteger = 1
		qInsere.ParamByName("MODULO").AsInteger = CInt(SessionVar("modulo"))
		qInsere.ParamByName("INICIOPROCESSO").AsDateTime = ServerNow
		StartTransaction
		qInsere.ExecSQL
		Commit


			On Error GoTo FIM:

				Dim qUpdate As Object
				Set qUpdate = NewQuery

				qUpdate.Add("UPDATE R_RELATORIOS_GERADOS SET CONCLUSAOPROCESSO =:CONCLUSAOPROCESSO, OCORRENCIA =:OCORRENCIA, SITUACAO =:SITUACAO, TAMANHOARQUIVO =:FILESIZE WHERE HANDLE = :HANDLE")
				qUpdate.ParamByName("HANDLE").AsInteger = vHandleInsercao

				Dim vCaminhoRelatorio As String
				Dim vCaminhoRelatorioU As String


				vCaminhoRelatorio = SessionPath  &  qHandleRelatorio.FieldByName("NOME").AsString & IIf(UserVar("TIPOARQUIVO")="",".pdf",UserVar("TIPOARQUIVO"))

				relatorio.ExportToFile(vCaminhoRelatorio)


				If UserVar("TIPOARQUIVO") = ".DAT" Then
				    RenameFile(vCaminhoRelatorio,SessionPath  &  qHandleRelatorio.FieldByName("NOME").AsString & ".txt")
					vCaminhoRelatorio = SessionPath  &  qHandleRelatorio.FieldByName("NOME").AsString & ".txt"
				    UserVar("TIPOARQUIVO") = UserVar("TIPOARQUIVOANTERIOR")
				End If

			    SetFieldDocument("R_RELATORIOS_GERADOS","RELATORIO",vHandleInsercao,vCaminhoRelatorio,True)
				qUpdate.ParamByName("OCORRENCIA").AsString = String_Filtro
				qUpdate.ParamByName("FILESIZE").AsFloat = Round(FileLen(vCaminhoRelatorio) / 1024,2)
				qUpdate.ParamByName("CONCLUSAOPROCESSO").AsDateTime = ServerNow
			    Kill(vCaminhoRelatorio)

				qUpdate.ParamByName("SITUACAO").AsInteger = 2
				StartTransaction
				qUpdate.ExecSQL
				Commit

				Call Avisar(qHandleRelatorio.FieldByName("NOME").AsString,"Gerado com sucesso.")

				Set qUpdate = Nothing
				Set qInsere = Nothing
				Set relatorio = Nothing

                Exit Sub


		    FIM:


				qUpdate.ParamByName("OCORRENCIA").AsString = String_Filtro & vbNewLine & "Erro ao emitir o relatório " & qHandleRelatorio.FieldByName("NOME").AsString & vbNewLine & Err.Description
				qUpdate.ParamByName("SITUACAO").AsInteger = 3
				StartTransaction
				qUpdate.ExecSQL
				Commit

				Call Avisar(qHandleRelatorio.FieldByName("NOME").AsString,Err.Description )

				Set qUpdate = Nothing
				Set qInsere = Nothing
				Set relatorio = Nothing



End Sub



Sub Avisar(relatorio As String, situacao As String)
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Add("SELECT EMAIL ")
  Sql.Add("  FROM Z_GRUPOUSUARIOS")
  Sql.Add(" WHERE HANDLE = :PHANDLE")
  Sql.ParamByName("PHANDLE").AsInteger = CurrentUser
  Sql.Active = True

  Dim Mail As Object
  Set Mail = NewMail

    Mail.Subject="**** AVISO DE GERAÇÃO DE RELATÓRIO ****"
    Mail.From  = Sql.FieldByName("EMAIL").AsString
    Mail.SendTo= Sql.FieldByName("EMAIL").AsString
    Mail.Text.Add("Situação da geração do relatório " & relatorio )
    Mail.Text.Add(situacao)
    Mail.Send

  Set Sql = Nothing
  Set Mail = Nothing

End Sub
