'HASH: F0A576DC4BC07D813650AD57E217A448
' atualizada em 10/08/2007
'#Uses "*bsShowMessage"

Dim Situacao As Integer

Public Sub GERAALERTAPRESTADOR_OnClick()
	Dim SPP As Object
	Set SPP = NewStoredProc

	SPP.AutoMode = True
	SPP.Name = "BSAEX_GERAALERTAPRESTADOR"

	SPP.AddParam("P_PRESTADOR", ptInput)        'Int

	SPP.ParamByName("P_PRESTADOR").DataType   = ftInteger

	SPP.AddParam("P_USUARIO",ptInput)           'Int

	SPP.ParamByName("P_USUARIO").DataType = ftInteger

	SPP.AddParam("P_HANDLETABELA",ptInput)           'Int

	SPP.ParamByName("P_HANDLETABELA").DataType = ftInteger

	SPP.AddParam("P_RETORNO",ptOutput)          'Varchar(100)

	SPP.ParamByName("P_RETORNO").DataType      = ftString
	SPP.ParamByName("P_HANDLETABELA").AsInteger  = CurrentQuery.FieldByName("HANDLE").AsInteger
	SPP.ParamByName("P_PRESTADOR").AsInteger  = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	SPP.ParamByName("P_USUARIO").AsInteger    = CurrentUser

	SPP.ExecProc

	If SPP.ParamByName("P_RETORNO").AsString <> "" Then
		bsShowMessage(SPP.ParamByName("P_RETORNO").AsString, "I")
	End If

	RefreshNodesWithTable("AEX_PRESTADORESCONECT")
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String
	Dim vDataHora As String

	' Luiz Gustavo - SMS 93073 - 04/03/2008 - Início
	'vDataHora = FormatDateTime2("DD/MM/YYYY",ServerDate)
	vDataHora = CurrentSystem.SQLDateTime(ServerDate)
	' 93073 - Fim


	If PRESTADOR.PopupCase <>0 Then
		ShowPopup = False

		Set interface = CreateBennerObject("Procura.Procurar")

		vCabecs = "Prestador|Nome|CPFCNPJ"
		vColunas = "PRESTADOR|NOME|CPFCNPJ"
		vCriterio = " RECEBEDOR = 'S' "
		vCriterio = vCriterio + " AND DATACREDENCIAMENTO IS NOT NULL AND DATADESCREDENCIAMENTO IS NULL "
		vCriterio = vCriterio + " AND NOT EXISTS (Select 1 FROM AEX_PRESTADORESCONECT X WHERE  X.DATAINICIAL <= " + vDataHora + " AND DATAFINAL IS NULL AND SAM_PRESTADOR.HANDLE = X.PRESTADOR)"
		
		vTabela = "SAM_PRESTADOR"
		vTitulo = "Prestadores"
		vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <> 0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("PRESTADOR").AsInteger    = vHandle
			CurrentQuery.FieldByName("CODPRESTADOR").AsInteger = vHandle
		End If

		Set interface = Nothing
	Else
	ShowPopup = True
	End If
End Sub

Public Sub TABLE_AfterPost()
	Dim QInsereEnd As Object
	Dim SQL1 As Object
	Dim QUpdate As Object
	Dim vHandle As Long
	Dim VPrestador As Long
	Dim vHandlePrestador As Long
	Dim vDataHora As Date

	If Situacao = 3 Then
		' ******** Inicio SMS - 39471 - 30/06/2005 - Drummond ********
		vDataHora = ServerDate

		Set SQL1 = NewQuery

		SQL1.Clear

		SQL1.Add("SELECT DISTINCT                                                ")
		SQL1.Add("       CASE")
		SQL1.Add("            WHEN ((Select A.HANDLE FROM SAM_PRESTADOR_AFASTAMENTO A")
		SQL1.Add("                WHERE A.PRESTADOR = P.HANDLE")
		SQL1.Add("                  And A.DATAINICIAL <= :pDataIni")
		SQL1.Add("                  And (A.DATAFINAL Is Null Or A.DATAFINAL >= :pDataFim))) is Null THEN 'A'")
		SQL1.Add("                Else 'I' END SITUACAOPRESTADOR,") '1 SITUACAOPRESTADOR
		SQL1.Add("       P.HANDLE PHANDLE,") '2 PHANDLE
		SQL1.Add("       PE.HANDLE PEHANDLE,") '3 PEHANDLE
		SQL1.Add("       CASE WHEN PE.DATACANCELAMENTO IS NULL THEN NULL ELSE PE.DATACANCELAMENTO END PEDATACANCELAMENTO,") '4 DATACANCELAMENTO
		SQL1.Add("       P.PRESTADOR PPRESTADOR") '5 PRESTADOR
		SQL1.Add("     FROM AEX_PRESTADORESCONECT      PC")
		SQL1.Add("     JOIN SAM_PRESTADOR              P  On (P.HANDLE = PC.PRESTADOR)")
		SQL1.Add("LEFT JOIN SAM_PRESTADOR_AFASTAMENTO  PA On (PA.PRESTADOR = P.HANDLE)")
		SQL1.Add("LEFT JOIN SAM_PRESTADOR_ENDERECO     PE On (PE.PRESTADOR = P.HANDLE)")
		SQL1.Add("    WHERE PC.DATAINICIAL <= :pDataIni")
		SQL1.Add("      And (PC.DATAFINAL Is Null Or PC.DATAFINAL >= :pDataFim)")
		SQL1.Add("      And P.HANDLE = :pHandle")

		SQL1.ParamByName("pDataIni").Value = ServerDate
		SQL1.ParamByName("pDataFim").Value = ServerDate
		SQL1.ParamByName("pHandle").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL1.Active = True

		SQL1.First

		Set QInsereEnd = NewQuery

		Do While Not SQL1.EOF

			QInsereEnd.Clear

			QInsereEnd.Add("INSERT INTO AEX_PRESTADOR_PRS(")
			QInsereEnd.Add("             HANDLE,            ") '1
			QInsereEnd.Add("             PRESTADORESCONECT, ") '18
			QInsereEnd.Add("             EMPCONECT,         ") '2
			QInsereEnd.Add("             SITUACAOPRESTADOR, ") '3
			QInsereEnd.Add("             HORARIOINICIOATEND,") '4
			QInsereEnd.Add("             HORARIOFIMATEND,   ") '5
			QInsereEnd.Add("             PRESTADOR,         ") '6
			QInsereEnd.Add("             CODPRESTADOR,      ") '7
			QInsereEnd.Add("             PERIODORETORNO,    ") '8
			QInsereEnd.Add("             TIPODEATENDIMENTO, ") '9
			QInsereEnd.Add("             TIPODERETORNO,     ") '10
			QInsereEnd.Add("             FORMADERETORNO,    ") '11
			QInsereEnd.Add("             SEQUENCIAENDERECO, ") '12
			QInsereEnd.Add("             PROCESSADO,        ") '13
			QInsereEnd.Add("             DATAINCLUSAO,      ") '14
			QInsereEnd.Add("             DATAALTERACAO,     ") '15
			QInsereEnd.Add("             USUARIO,           ") '16
			QInsereEnd.Add("             DATCANCENDERECO)   ") '17
			QInsereEnd.Add("VALUES (")
			QInsereEnd.Add("             :pNovoHandle,      ") '1
			QInsereEnd.Add("             :pHandle,          ") '18
			QInsereEnd.Add("             :pEMPCONECT,       ") '2
			QInsereEnd.Add("             :pSitPrestador,    ") '3
			QInsereEnd.Add("             Null,              ") '4
			QInsereEnd.Add("             Null,              ") '5
			QInsereEnd.Add("             :pPHandle,         ") '6
			QInsereEnd.Add("             :pPrestador,       ") '7
			QInsereEnd.Add("             Null,              ") '8
			QInsereEnd.Add("             3,                 ") '9
			QInsereEnd.Add("             3,                 ") '10
			QInsereEnd.Add("             7,                 ") '11
			QInsereEnd.Add("             :pPEHandle,        ") '12
			QInsereEnd.Add("             'N',               ") '13
			QInsereEnd.Add("             :pDataHora,        ") '14
			QInsereEnd.Add("             Null,              ") '15
			QInsereEnd.Add("             :pUsuario,         ") '16

			If SQL1.FieldByName("PEDATACANCELAMENTO").AsString = "" Then
				QInsereEnd.Add("         Null)") '17
			Else
				QInsereEnd.Add("         :pDataCancel)") '17
			End If

			'Parametros da Query SQL1
			QInsereEnd.ParamByName("pSitPrestador").Value = SQL1.FieldByName("SITUACAOPRESTADOR").AsString
			QInsereEnd.ParamByName("pPHandle").Value = SQL1.FieldByName("PHandle").AsInteger
			QInsereEnd.ParamByName("pPrestador").Value = SQL1.FieldByName("PPRESTADOR").AsString
			QInsereEnd.ParamByName("pPEHandle").Value = SQL1.FieldByName("PEHandle").AsInteger

			If SQL1.FieldByName("PEDATACANCELAMENTO").AsString <> "" Then
				QInsereEnd.ParamByName("pDataCancel").Value = SQL1.FieldByName("PEDATACANCELAMENTO").AsDateTime
			End If

			'Gera um novo handle para a tabela AEX_PRESTADOR_PRS
			QInsereEnd.ParamByName("pNOVOHANDLE").Value = NewHandle("AEX_PRESTADOR_PRS")
			QInsereEnd.ParamByName("pHandle").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
			QInsereEnd.ParamByName("pEMPCONECT").Value = CurrentQuery.FieldByName("EMPCONECT").AsInteger
			QInsereEnd.ParamByName("pDataHora").Value = vDataHora
			QInsereEnd.ParamByName("pUsuario").Value = CurrentUser

			QInsereEnd.ExecSQL
			SQL1.Next

		Loop

		Set QUpdate = NewQuery

		QUpdate.Clear

		QUpdate.Add("UPDATE AEX_PRESTADOR_PRS")
		QUpdate.Add("     Set SITUACAOPRESTADOR = 'I'")
		QUpdate.Add("   WHERE DATCANCENDERECO Is Not Null")
		QUpdate.Add("     AND PRESTADOR = :pPRESTADOR") 'Fim do UPDATE

		QUpdate.ParamByName("pPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger

		QUpdate.ExecSQL

		' ******** Fim SMS - 39471 - 30/06/2005 - Drummond ********
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim interface As Object
	Dim Linha As String
	Dim Condicao As String
	Dim qverifica As Object
	Set qverifica = NewQuery

	qverifica.Add("SELECT HANDLE ")
	qverifica.Add("  FROM AEX_PRESTADORESCONECT")
	qverifica.Add(" WHERE PRESTADOR = :PRESTADOR")
	qverifica.Add("AND EMPCONECT = :EMPCON")
	qverifica.Add("AND ((DATAFINAL > :DATAFINAL) OR (DATAFINAL IS NULL))")
	qverifica.Add("AND DATAINICIAL <= :DATAINICIAL")
	qverifica.Add("AND HANDLE <> :HANDLE")

	qverifica.ParamByName("EMPCON").Value = CurrentQuery.FieldByName("EMPCONECT").AsInteger
	qverifica.ParamByName("DATAFINAL").Value = ServerDate
	qverifica.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	qverifica.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	qverifica.ParamByName("DATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime

	qverifica.Active = True

	If Not qverifica.EOF Then
		bsShowMessage("Este prestador tem uma vigência com data Final aberta!", "E")
		CanContinue = False
	End If

	qverifica.Active = False

	CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

	If (CanContinue = True ) And (CurrentQuery.State = 3) Then
		Situacao = CurrentQuery.State
	End If
End Sub

Public Sub TABLE_NewRecord()
	CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "GERAALERTAPRESTADOR"
			GERAALERTAPRESTADOR_OnClick
	End Select
End Sub
