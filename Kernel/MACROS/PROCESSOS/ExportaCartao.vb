'HASH: 1148D03FABD6E3FC3AAFF833E4591EFD
Option Explicit

Public Sub BOTAOEXPORTAR_OnClick(CanInherited As Boolean)
	CanInherited = False

	Dim QueryIntervalo As Object
	Set QueryIntervalo = NewQuery

	Dim QueryUsuario As Object
	Set QueryUsuario = NewQuery

	Dim	QueryRelatorio As Object
	Set QueryRelatorio = NewQuery

	Dim	QueryOcorrencia As Object
	Set QueryOcorrencia = NewQuery

	Dim vHandleInicial As Long
	Dim VContinue As Boolean
	Dim vSequenciaArquivo As Integer
	Dim vExtensao As String

	vHandleInicial = 0
	VContinue = True
	vSequenciaArquivo = 1


	QueryOcorrencia.Add("UPDATE SAM_ROTINACARTAO")
	QueryOcorrencia.Add("SET OCORRENCIAS = OCORRENCIAS + :OCORRENCIAS")
	QueryOcorrencia.Add("WHERE HANDLE =:HANDLE")


	QueryIntervalo.Add("SELECT MIN(HANDLE) INICIAL, MAX(HANDLE) FINAL, COUNT(HANDLE) QTDE")
	QueryIntervalo.Add("FROM (SELECT A.HANDLE")
	QueryIntervalo.Add("        FROM SAM_BENEFICIARIO_CARTAOIDENTIF A")
	QueryIntervalo.Add("        JOIN SAM_ROTINACARTAO_CARTAO B ON (A.HANDLE = B.CARTAOIDENTIFICACAO)")
	QueryIntervalo.Add("        JOIN SAM_ROTINACARTAO C ON (C.HANDLE = B.ROTINACARTAO)")
	QueryIntervalo.Add("       WHERE A.HANDLE > :HANDLEINICIAL AND A.SITUACAO <> 'C'")
	QueryIntervalo.Add("         AND C.HANDLE =:PROTINACARTAO")
	QueryIntervalo.Add("       ORDER BY A.HANDLE) WHERE ROWNUM <=3332")


	QueryRelatorio.Add("SELECT RELATORIOCARTAO FROM SAM_PARAMETROSBENEFICIARIO")
	QueryRelatorio.Active = True

	If QueryRelatorio.FieldByName("RELATORIOCARTAO").IsNull Then
		If VisibleMode Then
			MsgBox("Parametrize o Relatório utilizado para exportação. Parâmetros do Beneficiário, página Cartão.")
		Else
			CancelDescription = "Parametrize o Relatório utilizado para exportação. Parâmetros do Beneficiário, página Cartão."
		End If
	Else
		If Not CurrentQuery.FieldByName("ARQUIVOCONTRATO").IsNull Then

			QueryUsuario.Add("UPDATE SAM_ROTINACARTAO SET USUARIOEXPORTACAO =:USUARIOEXPORTACAO, DATAEXPORTACAO =:DATAEXPORTACAO WHERE HANDLE =:HANDLE")
			QueryUsuario.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			QueryUsuario.ParamByName("USUARIOEXPORTACAO").AsInteger = CurrentUser
			QueryUsuario.ParamByName("DATAEXPORTACAO").AsDateTime = ServerNow
			QueryUsuario.ExecSQL

			QueryOcorrencia.Active = False
			QueryOcorrencia.ParamByName("HANDLE").AsInteger       = NewHandle("SAM_ROTINACARTAO_OCORRENCIA")
			QueryOcorrencia.ParamByName("OCORRENCIAS").AsString   = Chr(13)+"Iniciando a Exportação dos cartões "+FormatDateTime2("dd/mm/yyyy hh:mm",ServerNow)
			QueryOcorrencia.ExecSQL

			If InTransaction Then Commit

			Do While VContinue
				QueryIntervalo.Active = False
				QueryIntervalo.ParamByName("HANDLEINICIAL").AsInteger = vHandleInicial
				QueryIntervalo.ParamByName("PROTINACARTAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
				QueryIntervalo.Active = True

				VContinue = QueryIntervalo.FieldByName("QTDE").AsInteger = 3332

				vHandleInicial = QueryIntervalo.FieldByName("FINAL").AsInteger + 1 'para ser utilizado no proximo select
				UserVar("HANDLEINICIAL") = QueryIntervalo.FieldByName("INICIAL").AsInteger
				UserVar("HANDLEFINAL")   = QueryIntervalo.FieldByName("FINAL").AsInteger

				vExtensao = ".DAT"


				If InStr(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString,".DAT") = 0 Then
     				vExtensao = Mid(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString,Len(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString)-3,4)
     			End If

		  		ReportExport(QueryRelatorio.FieldByName("RELATORIOCARTAO").AsInteger, _
		  		"A.HANDLE = "+CurrentQuery.FieldByName("HANDLE").AsString, _
		  		Mid(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString,1,Len(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString)-4) _
		  		+"_"+Trim(Str(vSequenciaArquivo))+".DAT", _
		  		True,False)

		  		RenameFile(Mid(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString,1,Len(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString)-4)+"_"+Trim(Str(vSequenciaArquivo))+".DAT", _
				Mid(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString,1,Len(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString)-4) _
		  		+"_"+Trim(Str(vSequenciaArquivo))+vExtensao)

				If Not InTransaction Then StartTransaction

				QueryOcorrencia.Active = False
				QueryOcorrencia.ParamByName("HANDLE").AsInteger       = CurrentQuery.FieldByName("HANDLE").AsInteger
				QueryOcorrencia.ParamByName("OCORRENCIAS").AsString   = Chr(13)+"Gerado o arquivo "+Mid(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString,1,Len(CurrentQuery.FieldByName("ARQUIVOCONTRATO").AsString)-4)+"_"+Str(vSequenciaArquivo)+vExtensao+" com "+Str(QueryIntervalo.FieldByName("QTDE").AsInteger*3+2)+" linhas"
				QueryOcorrencia.ExecSQL

				If InTransaction Then Commit


		  		vSequenciaArquivo = vSequenciaArquivo + 1
			Loop

			If Not InTransaction Then StartTransaction


			QueryOcorrencia.Active = False
			QueryOcorrencia.ParamByName("HANDLE").AsInteger       = CurrentQuery.FieldByName("HANDLE").AsInteger
			QueryOcorrencia.ParamByName("OCORRENCIAS").AsString   = Chr(13)+"Finalizando a exportação dos cartões "+FormatDateTime2("dd/mm/yyyy hh:mm",ServerNow)
			QueryOcorrencia.ExecSQL


			If InTransaction Then Commit


		Else
			If VisibleMode Then
				MsgBox("Informe o caminho de destino")
				ARQUIVOCONTRATO.SetFocus
			Else
				CancelDescription = "Informe o caminho de destino"
			End If
		End If
	End If

	Set QueryRelatorio = Nothing
	Set QueryIntervalo = Nothing
	Set QueryOcorrencia = Nothing
	Set QueryUsuario = Nothing
End Sub
