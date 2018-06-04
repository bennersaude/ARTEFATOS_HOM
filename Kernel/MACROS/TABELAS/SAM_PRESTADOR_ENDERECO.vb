'HASH: F7D3AF2A3BBDE5F8C61323E73477CF2B
'Macro: SAM_PRESTADOR_ENDERECO
'#Uses "*bsShowMessage"
Option Explicit

Dim vgDataCancelamento As Date
Dim gbAtendimento As Boolean

Public Sub ATENDIMENTO_OnChange()
	gbAtendimento = Not gbAtendimento
	CNES.Visible = gbAtendimento
End Sub

Public Sub CEP_OnChange()
  If (Len(CurrentQuery.FieldByName("CEP").AsString) < 9) Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ESTADO").Clear
    CurrentQuery.FieldByName("MUNICIPIO").Clear
    CurrentQuery.FieldByName("BAIRRO").Clear
    CurrentQuery.FieldByName("LOGRADOURO").Clear
    CurrentQuery.FieldByName("TIPOLOGRADOURO").Clear
    CurrentQuery.FieldByName("COMPLEMENTO").Clear
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO, TIPOLOGRADOURO,COMPLEMENTO   ")
  SQL.Add("  FROM LOGRADOUROS      ")
  SQL.Add(" WHERE CEP = :HANDLE ")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CEP").AsString
  SQL.Active = True

  If SQL.FieldByName("CEP").IsNull Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ESTADO").Clear
    CurrentQuery.FieldByName("MUNICIPIO").Clear
    CurrentQuery.FieldByName("BAIRRO").Clear
    CurrentQuery.FieldByName("LOGRADOURO").Clear
    CurrentQuery.FieldByName("TIPOLOGRADOURO").Clear
    CurrentQuery.FieldByName("COMPLEMENTO").Clear
  Else
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ESTADO").Value      = SQL.FieldByName("ESTADO").AsString
    CurrentQuery.FieldByName("MUNICIPIO").Value   = SQL.FieldByName("MUNICIPIO").AsString
    CurrentQuery.FieldByName("BAIRRO").Value      = SQL.FieldByName("BAIRRO").AsString
    CurrentQuery.FieldByName("LOGRADOURO").Value  = SQL.FieldByName("LOGRADOURO").AsString

	If (SQL.FieldByName("TIPOLOGRADOURO").IsNull) Then

	    Dim SQL2 As Object
	    Set SQL2 = NewQuery

        SQL2.Add("SELECT HANDLE           ")
        SQL2.Add("  FROM LOGRADOUROS_TIPO ")
        SQL2.Add(" WHERE CODIGO = '081'   ")
        SQL2.Active = True

	    SessionVar("TIPOLOGRADOURO") = "0"  ' Paulo Melo - SMS 135779 - 13/05/2010
		CurrentQuery.FieldByName("TIPOLOGRADOURO").Value = SQL2.FieldByName("HANDLE").AsInteger
		Set SQL2 = Nothing

	Else
		SessionVar("TIPOLOGRADOURO") = SQL.FieldByName("TIPOLOGRADOURO").AsString  ' Paulo Melo - SMS 135779 - 13/05/2010
		CurrentQuery.FieldByName("TIPOLOGRADOURO").Value = SQL.FieldByName("TIPOLOGRADOURO").AsString
	End If
    CurrentQuery.FieldByName("COMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString
  End If
  Set SQL = Nothing
End Sub

Public Sub CEP_OnPopup(ShowPopup As Boolean)
	' Joldemar Moreira 18/08/2003
	' SMS 16059
	Dim vHandle As String
  Dim interface As Object
  ShowPopup = False
  Set interface = CreateBennerObject("ProcuraCEP.Rotinas")
  interface.Exec(CurrentSystem, vHandle)

  If vHandle <>"" Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO, TIPOLOGRADOURO,COMPLEMENTO   ")
    SQL.Add("  FROM LOGRADOUROS      ")
    SQL.Add(" WHERE CEP = :HANDLE ")
    SQL.ParamByName("HANDLE").Value = vHandle
    SQL.Active = True

    CurrentQuery.Edit
    CurrentQuery.FieldByName("CEP").Value = SQL.FieldByName("CEP").AsString
    CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
    CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
    CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
    CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString

	If (SQL.FieldByName("TIPOLOGRADOURO").IsNull) Then
		SessionVar("TIPOLOGRADOURO") = "0"  ' Paulo Melo - SMS 135779 - 13/05/2010
		CurrentQuery.FieldByName("TIPOLOGRADOURO").Value = 81
	Else
		SessionVar("TIPOLOGRADOURO") = SQL.FieldByName("TIPOLOGRADOURO").AsString  ' Paulo Melo - SMS 135779 - 13/05/2010
		CurrentQuery.FieldByName("TIPOLOGRADOURO").Value = SQL.FieldByName("TIPOLOGRADOURO").AsString
	End If
    CurrentQuery.FieldByName("COMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString

  End If

  Set interface = Nothing
End Sub

Public Sub MOTIVOCANCELAMENTO_OnChange()
	If(CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull)Then
		CurrentQuery.FieldByName("DATACANCELAMENTO").Value = Null
	Else
		CurrentQuery.FieldByName("DATACANCELAMENTO").Value = ServerDate
	End If
End Sub

Public Sub MOTIVOCANCELAMENTO_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String

	ShowPopup = False

	Set interface = CreateBennerObject("Procura.Procurar")

	vTabela = "SAM_MOTIVOCANCELAMENTOENDERECO"
	vColunas = "SAM_MOTIVOCANCELAMENTOENDERECO.CODIGO|SAM_MOTIVOCANCELAMENTOENDERECO.MOTIVOCANCELAMENTO"
	vCriterio = "SAM_MOTIVOCANCELAMENTOENDERECO.CODIGO > 0"
	vCampos = "Código|Motivo"
	vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "Motivo de Cancelamento de Endereço", True, "")

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsInteger = vHandle
		CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime = ServerDate
	End If

	Set interface = Nothing
End Sub

Public Sub TABLE_AfterCommitted()
	If (CurrentQuery.FieldByName("CORRESPONDENCIA").Value = "S") Then
		Dim SamPrestadorBLL As CSBusinessComponent

	    Set SamPrestadorBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.SamPrestadorBLL, Benner.Saude.Prestadores.Business")
	   	SamPrestadorBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
		SamPrestadorBLL.Execute("VerificarSeExportaBennerHospitalar")
	End If
End Sub

Public Sub TABLE_AfterInsert()
	gbAtendimento = CurrentQuery.FieldByName("ATENDIMENTO").AsString = "S"
	CNES.Visible = gbAtendimento
End Sub

Public Sub TABLE_AfterPost()
	If(CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull)Then
		Componentes_ReadOnly(False)
	Else
		Componentes_ReadOnly(True)
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If(CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull)Then
		Componentes_ReadOnly(False)
	Else
		Componentes_ReadOnly(True)
	End If

	gbAtendimento = CurrentQuery.FieldByName("ATENDIMENTO").AsString = "S"
	CNES.Visible = gbAtendimento
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Msg As String
	Dim sqlRecuperaRegistro As Object

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage( Msg,"E")
		CanContinue = False
		Exit Sub
	End If

	Dim SQL As Object
	If CurrentQuery.FieldByName("CORRESPONDENCIA").Value = "S" Then
		Set SQL = NewQuery

		SQL.Add("SELECT CORRESPONDENCIA FROM SAM_PRESTADOR_ENDERECO    ")
		SQL.Add("  WHERE PRESTADOR = :PRESTADOR                        ")
		SQL.Add("    AND CORRESPONDENCIA = 'S' AND HANDLE <> :HCORRENTE")
		SQL.Add("    AND DATACANCELAMENTO IS NULL                      ")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
		SQL.ParamByName("HCORRENTE").Value = CurrentQuery.FieldByName("HANDLE").Value
		SQL.Active = True

		If Not SQL.EOF Then
			CurrentQuery.FieldByName("CORRESPONDENCIA").Value = "N"
			CanContinue = False
			bsShowMessage( "Existe outro endereco marcado para correspondência!","E")
			Exit Sub
		End If

		Set SQL = Nothing
	End If

	If CurrentQuery.FieldByName("PESSOAL").Value = "S" Then
		Set SQL = NewQuery

		SQL.Add("SELECT PESSOAL FROM SAM_PRESTADOR_ENDERECO            ")
		SQL.Add("  WHERE PRESTADOR = :PRESTADOR                        ")
		SQL.Add("    AND PESSOAL = 'S' AND HANDLE <> :HCORRENTE        ")
		SQL.Add("    AND DATACANCELAMENTO IS NULL                      ")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
		SQL.ParamByName("HCORRENTE").Value = CurrentQuery.FieldByName("HANDLE").Value
		SQL.Active = True

		If Not SQL.EOF Then
			CurrentQuery.FieldByName("PESSOAL").Value = "N"
			CanContinue = False
			bsShowMessage( "Existe outro endereco marcado como pessoal!","E")
			Exit Sub
		End If

		Set SQL = Nothing
	End If


	If CurrentQuery.FieldByName("CORRESPONDENCIA").Value = "N" And CurrentQuery.FieldByName("ATENDIMENTO").Value = "N" And CurrentQuery.FieldByName("PESSOAL").Value = "N" Then
		bsShowMessage( "Endereço deve ser marcado para correspondência, atendimento ou pessoal!","E")
		CanContinue = False
		Exit Sub
	End If

	If VerificaEnderecoExistente Then
		bsShowMessage("Já existe este endereço para este prestador." ,"E")
		CanContinue = False
		Exit Sub
	End If

	If CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull Then
		CurrentQuery.FieldByName("DATACANCELAMENTO").Value = Null
	Else
		CurrentQuery.FieldByName("DATACANCELAMENTO").Value = ServerDate

		Dim EnderecoLivro As Object
		Set EnderecoLivro = NewQuery

		EnderecoLivro.Add("SELECT ENDERECO FROM SAM_PRESTADOR_LIVRO WHERE PRESTADOR = :PRESTADOR")

		EnderecoLivro.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		EnderecoLivro.Active = True

		EnderecoLivro.First

		While(Not EnderecoLivro.EOF)
			If(EnderecoLivro.FieldByName("ENDERECO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger)Then
				bsShowMessage("Este endereço esta cadastrado no livro, não pode ser cancelado." ,"E")
				CanContinue = False
				Exit Sub
			End If

			EnderecoLivro.Next
		Wend

		Set EnderecoLivro = Nothing
	End If


	If CurrentQuery.FieldByName("ATENDIMENTO").Value = "N" Then
		CurrentQuery.FieldByName("CNES").Clear
	Else
		If CurrentQuery.FieldByName("CNES").IsNull Then
			Dim vsCNES As String
			vsCNES = "Cadastro Nacional de Estabelec. de Saúde (CNES) não foi preenchido. Como trata-se de um endereço de atendimento o CNES é considerado obrigatório pela ANS."

			If VisibleMode Then 'SMS 119363/119393 - Ricardo Rocha - 17/08/2009
				If bsShowMessage(vsCNES + "Deseja continuar da mesma forma?", "Q") = vbNo Then
					CNES.SetFocus
					CanContinue = False
					Exit Sub
				End If
			Else
				bsShowMessage(vsCNES + Chr(13) + "Para incluí-lo será necessário editar o registro digitar um CNES.", "E")
			End If
 		End If
	End If


	Dim vHandleEndereci  As Long
	Dim VerificaConect   As Object
	Dim AlteraEndereco   As Object
	Dim InsereEndereco   As Object
	Dim VerificaEndereco As Object
	Dim vCodPrestador    As String
	Dim BuscaPrestador   As Object
	Dim vHandle			As Long
	Set VerificaConect = NewQuery

	VerificaConect.Add("SELECT ATUALIZAENDERECOPRESTADOR FROM AEX_PARAMETROSGERAIS")

	VerificaConect.Active = True

	If VerificaConect.FieldByName("ATUALIZAENDERECOPRESTADOR").AsString = "S" Then
		Set SQL = NewQuery

		SQL.Clear

		SQL.Active = False

		SQL.Add("SELECT HANDLE FROM AEX_PRESTADORESCONECT WHERE PRESTADOR = :PRESTADOR")

		SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
		SQL.Active = True

		If SQL.FieldByName("HANDLE").AsInteger > 0 Then
			Set VerificaEndereco = NewQuery

			VerificaEndereco.Add("SELECT HANDLE                                  ")
			VerificaEndereco.Add("  FROM AEX_PRESTADOR_PRS                       ")
			VerificaEndereco.Add(" WHERE PRESTADOR = :PRESTADOR                  ")
			VerificaEndereco.Add("   AND PRESTADORESCONECT = :PRESTADORESCONECT  ")
			VerificaEndereco.Add("   AND SEQUENCIAENDERECO = :SEQUENCIAENDERECO  ")

			VerificaEndereco.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
			VerificaEndereco.ParamByName("PRESTADORESCONECT").Value = SQL.FieldByName("HANDLE").AsInteger
			VerificaEndereco.ParamByName("SEQUENCIAENDERECO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
			VerificaEndereco.Active = True

			If VerificaEndereco.FieldByName("HANDLE").AsInteger > 0 Then
				StartTransaction

				Set AlteraEndereco = NewQuery

				AlteraEndereco.Add("UPDATE AEX_PRESTADOR_PRS                  ")
				AlteraEndereco.Add("   SET DATAALTERACAO   = :DATAALTERACAO,  ")
				AlteraEndereco.Add("       PROCESSADO      = :PROCESSADO,     ")
				AlteraEndereco.Add("       USUARIO         = :USUARIO,        ")
				AlteraEndereco.Add("       DATCANCENDERECO = :DATCANCENDERECO ")
				AlteraEndereco.Add(" WHERE HANDLE = :HANDLE                   ")

				AlteraEndereco.ParamByName("DATAALTERACAO").Value   = ServerNow
				AlteraEndereco.ParamByName("PROCESSADO").Value      = "N"
				AlteraEndereco.ParamByName("USUARIO").Value         = CurrentUser
				AlteraEndereco.ParamByName("DATCANCENDERECO").Value = ServerDate
				AlteraEndereco.ParamByName("HANDLE").Value          = VerificaEndereco.FieldByName("HANDLE").AsInteger

				AlteraEndereco.ExecSQL

				Commit
			Else
				Set VerificaConect = NewQuery

				StartTransaction

				Set InsereEndereco = NewQuery
				Set BuscaPrestador = NewQuery

				BuscaPrestador.Add("SELECT PRESTADOR FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")

				BuscaPrestador.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
				BuscaPrestador.Active = True

				vHandle = NewHandle("AEX_PRESTADOR_PRS")

				InsereEndereco.Add("INSERT INTO AEX_PRESTADOR_PRS   ")
				InsereEndereco.Add("           (HANDLE,             ")
				InsereEndereco.Add("            PRESTADORESCONECT,  ")
				InsereEndereco.Add("            EMPCONECT,          ")
				InsereEndereco.Add("            SITUACAOPRESTADOR,  ")
				InsereEndereco.Add("            PRESTADOR,          ")
				InsereEndereco.Add("            CODPRESTADOR,       ")
				InsereEndereco.Add("            TIPODEATENDIMENTO,  ")
				InsereEndereco.Add("            TIPODERETORNO,      ")
				InsereEndereco.Add("            FORMADERETORNO,     ")
				InsereEndereco.Add("            SEQUENCIAENDERECO,  ")
				InsereEndereco.Add("            PROCESSADO,         ")
				InsereEndereco.Add("            DATAINCLUSAO,       ")
				InsereEndereco.Add("            DATAALTERACAO,      ")
				InsereEndereco.Add("            USUARIO,            ")
				InsereEndereco.Add("            DATCANCENDERECO)    ")
				InsereEndereco.Add("            VALUES              ")
				InsereEndereco.Add("           (:HANDLE,            ")
				InsereEndereco.Add("            :PRESTADORESCONECT, ")
				InsereEndereco.Add("            :EMPCONECT,         ")
				InsereEndereco.Add("            :SITUACAOPRESTADOR, ")
				InsereEndereco.Add("            :PRESTADOR,         ")
				InsereEndereco.Add("            :CODPRESTADOR,      ")
				InsereEndereco.Add("            :TIPODEATENDIMENTO, ")
				InsereEndereco.Add("            :TIPODERETORNO,     ")
				InsereEndereco.Add("            :FORMADERETORNO,    ")
				InsereEndereco.Add("            :SEQUENCIAENDERECO, ")
				InsereEndereco.Add("            :PROCESSADO,        ")
				InsereEndereco.Add("            :DATAINCLUSAO,      ")
				InsereEndereco.Add("            :DATAALTERACAO,     ")
				InsereEndereco.Add("            :USUARIO,           ")

				InsereEndereco.ParamByName("HANDLE").Value            = vHandle
				InsereEndereco.ParamByName("PRESTADORESCONECT").Value = SQL.FieldByName("HANDLE").AsInteger
				InsereEndereco.ParamByName("EMPCONECT").Value         = 1
				InsereEndereco.ParamByName("SITUACAOPRESTADOR").Value = "A"
				InsereEndereco.ParamByName("PRESTADOR").Value         = CurrentQuery.FieldByName("PRESTADOR").AsInteger
				InsereEndereco.ParamByName("CODPRESTADOR").Value      = BuscaPrestador.FieldByName("PRESTADOR").AsString
				InsereEndereco.ParamByName("TIPODEATENDIMENTO").Value = 3
				InsereEndereco.ParamByName("TIPODERETORNO").Value     = 3
				InsereEndereco.ParamByName("FORMADERETORNO").Value    = 7
				InsereEndereco.ParamByName("SEQUENCIAENDERECO").Value = vHandle
				InsereEndereco.ParamByName("PROCESSADO").Value        = "N"
				InsereEndereco.ParamByName("DATAINCLUSAO").Value      = ServerNow
				InsereEndereco.ParamByName("DATAALTERACAO").Value     = ServerNow
				InsereEndereco.ParamByName("USUARIO").Value           = CurrentUser

				If  CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
					InsereEndereco.Add("            NULL)   ")
				Else
					InsereEndereco.Add("            :DATCANCENDERECO)   ")
					InsereEndereco.ParamByName("DATCANCENDERECO").Value   = CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime
				End If

				InsereEndereco.ExecSQL

				Commit
			End If
		End If

		VerificaConect.Active = False
	End If

  Set sqlRecuperaRegistro = NewQuery
  sqlRecuperaRegistro.Clear
  sqlRecuperaRegistro.Active = False
  sqlRecuperaRegistro.Add("SELECT LATITUDE, LONGITUDE FROM SAM_PRESTADOR_ENDERECO WHERE HANDLE = :HANDLE")
  sqlRecuperaRegistro.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_ENDERECO")
  sqlRecuperaRegistro.Active = True

  If CurrentQuery.FieldByName("LATITUDE").IsNull = False Or CurrentQuery.FieldByName("LONGITUDE").IsNull = False Then
    If CurrentQuery.FieldByName("LATITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LATITUDE").IsNull Or CurrentQuery.FieldByName("LONGITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LONGITUDE").IsNull Then
      CurrentQuery.FieldByName("DTATUALIZACAOLATITUDELONGITUDE").AsDateTime = CurrentSystem.ServerDate
    End If
    If CurrentQuery.FieldByName("LATITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LATITUDE").AsFloat Or CurrentQuery.FieldByName("LONGITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LONGITUDE").AsFloat Then
      CurrentQuery.FieldByName("DTATUALIZACAOLATITUDELONGITUDE").AsDateTime = CurrentSystem.ServerDate
    End If
  End If
  Set sqlRecuperaRegistro = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage( Msg,"E")
		CanContinue = False
		Exit Sub
	End If

	If(Not CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").IsNull)Then
		bsShowMessage( "Endereço de Prestador cancelado, não pode ser excluído","E")
		CanContinue = False
		Exit Sub
	End If

	Dim EnderecoLivro As Object
	Set EnderecoLivro = NewQuery

	EnderecoLivro.Add("SELECT ENDERECO ")
	EnderecoLivro.Add("FROM SAM_PRESTADOR_LIVRO ")
	EnderecoLivro.Add("WHERE PRESTADOR = :PRESTADOR ")

	EnderecoLivro.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	EnderecoLivro.Active = True

	EnderecoLivro.First

	While(Not EnderecoLivro.EOF)
		If(EnderecoLivro.FieldByName("ENDERECO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger)Then
			bsShowMessage("Este endereço esta cadastrado no livro, não pode ser excluido.", "E")
			CanContinue = False
			Exit Sub
		End If

		EnderecoLivro.Next
	Wend

	Set EnderecoLivro = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage( Msg,"E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub Componentes_ReadOnly(pbEnable As Boolean)
	ATENDIMENTO.ReadOnly = pbEnable
	BAIRRO.ReadOnly = pbEnable
	COMPLEMENTO.ReadOnly = pbEnable
	CORRESPONDENCIA.ReadOnly = pbEnable
	ESTADO.ReadOnly = pbEnable
	FAX.ReadOnly = pbEnable
	LOGRADOURO.ReadOnly = pbEnable
	TIPOLOGRADOURO.ReadOnly = pbEnable
	MUNICIPIO.ReadOnly = pbEnable
	NUMERO.ReadOnly = pbEnable
	PONTOREFERENCIA.ReadOnly = pbEnable
	PRESTADOR.ReadOnly = pbEnable
	QTDVAGASESTACIONAMENTO.ReadOnly = pbEnable
	CNES.ReadOnly = pbEnable
	RAMAL1.ReadOnly = pbEnable
	RAMAL2.ReadOnly = pbEnable
	TELEFONE1.ReadOnly = pbEnable
	TELEFONE2.ReadOnly = pbEnable
	CEP.Visible = Not pbEnable
	PESSOAL.ReadOnly = pbEnable
End Sub

Public Function VerificaEnderecoExistente As Boolean
	Dim vLogradouro As String
	Dim vTipoLogradouro As String
	Dim vComplemento As String
	Dim qVerifica As Object
	Set qVerifica = NewQuery

	VerificaEnderecoExistente = False

	qVerifica.Add("SELECT HANDLE, LOGRADOURO, TIPOLOGRADOURO, COMPLEMENTO")
	qVerifica.Add("  FROM SAM_PRESTADOR_ENDERECO                         ")
	qVerifica.Add(" WHERE HANDLE <> :HANDLE                              ")
	qVerifica.Add("   AND PRESTADOR = :PRESTADOR                         ")
	qVerifica.Add("   AND DATACANCELAMENTO IS NULL                       ")

	If Not CurrentQuery.FieldByName("NUMERO").IsNull Then
		qVerifica.Add("   AND NUMERO = :NUMERO")
		qVerifica.ParamByName("NUMERO").Value = CurrentQuery.FieldByName("NUMERO").AsInteger
	Else
		qVerifica.Add("   AND NUMERO IS NULL")
	End If

	qVerifica.Add("   AND CEP = :CEP")

	qVerifica.ParamByName("CEP").Value = CurrentQuery.FieldByName("CEP").AsString
	qVerifica.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	qVerifica.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	qVerifica.Active = True

	While(Not qVerifica.EOF)And(Not VerificaEnderecoExistente)
		vLogradouro = Replace(qVerifica.FieldByName("LOGRADOURO").AsString, " ", "")
		vTipoLogradouro = Replace(qVerifica.FieldByName("TIPOLOGRADOURO").AsString, " ", "")
		vComplemento = Replace(qVerifica.FieldByName("COMPLEMENTO").AsString, " ", "")

		If (vLogradouro = Replace(CurrentQuery.FieldByName("LOGRADOURO").AsString, " ", ""))And _
			(vTipoLogradouro = Replace(CurrentQuery.FieldByName("TIPOLOGRADOURO").AsString, " ", ""))And _
		   (vComplemento = Replace(CurrentQuery.FieldByName("COMPLEMENTO").AsString, " ", ""))Then
			VerificaEnderecoExistente = True
		End If

		qVerifica.Next
	Wend

	qVerifica.Active = False

	Set qVerifica = Nothing
End Function
