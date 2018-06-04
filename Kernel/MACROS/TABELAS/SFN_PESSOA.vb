'HASH: 72D170B2987BC097D83AD26B891A4144
'Macro: SFN_PESSOA
'Juliano -23/11/2000 -grava regiao e filialcusto
'#uses "*FormatarTelefone"
'#uses "*bsShowMessage"

Dim vsModoEdicao            As String
Dim vsXMLContainerEnderecos As String
Dim vsXMLEnderecosExcluidos As String

Option Explicit

Public Function CriaUsuario(pNome As String, pGrupo As String, pEmail As String, pLogin As String, pFilial As String) As Long
  Dim vUser As CSSystemUser
  Dim viHandle As Long
  Dim vbIncluiu As Boolean
  Dim vGrupo As Long


  Dim SQL As Object
  Set SQL = NewQuery

  On Error GoTo Erro

  CriaUsuario = 0

  vbIncluiu = False

  SQL.Clear
  SQL.Add("SELECT GRUPOSEGURANCAPESSOA FROM SAM_PARAMETROSWEB")
  SQL.Active = True

  vGrupo = SQL.FieldByName("GRUPOSEGURANCAPESSOA").AsInteger

  SQL.Clear
  SQL.Add("SELECT A.HANDLE, B.MATRICULAUNICA, C.PRESTADOR, D.PESSOA FROM Z_GRUPOUSUARIOS A")
  SQL.Add(" LEFT JOIN Z_GRUPOUSUARIOS_BENEFICIARIO B ON (B.USUARIO = A.HANDLE)")
  SQL.Add(" LEFT JOIN Z_GRUPOUSUARIOS_PRESTADOR    C ON (C.USUARIO = A.HANDLE)")
  SQL.Add(" LEFT JOIN Z_GRUPOUSUARIOS_PESSOA       D ON (D.USUARIO = A.HANDLE)")
  SQL.Add("WHERE APELIDO = :APELIDO")
  SQL.ParamByName("APELIDO").AsString = pLogin
  SQL.Active = True

  If SQL.FieldByName("HANDLE").AsInteger = 0 Then

    SQL.Clear
    SQL.Add("SELECT SMTPSERVER, SMTPPORT FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HANDLE ")
    SQL.ParamByName("HANDLE").AsInteger = CurrentUser
    SQL.Active = True

    Set vUser = NewSystemUser
    vUser.NewUser
    vUser.UserProperty("NOME") = pNome
    vUser.UserProperty("GRUPO") = pGrupo
    vUser.UserProperty("EMAIL") = pEmail
    vUser.UserProperty("APELIDO") = pLogin
    vUser.UserProperty("SENHA") = pLogin '"123456"
    vUser.UserProperty("FILIALPADRAO") = pFilial
    vUser.UserProperty("CODIGO") = Str(NewHandle("Z_GRUPOUSUARIOS"))
    vUser.UserProperty("SMTPSERVER") = SQL.FieldByName("SMTPSERVER").AsString
    vUser.UserProperty("SMTPPORT") = SQL.FieldByName("SMTPPORT").AsString
    vUser.UserProperty("ALTERARSENHA") = "S"


    ' Só abre transação se for criar o usuário de fato
    StartTransaction

    viHandle = vUser.SaveUserProperties
    vUser.SelectUser(viHandle)

    SQL.Clear
    SQL.Add("SELECT NOME, URL FROM EMPRESAS WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = CurrentCompany
    SQL.Active = True

    Dim vEmpresa As String
    Dim url As String

    Dim vCorpoTexto As String

    vEmpresa = SQL.FieldByName("NOME").AsString
    url = SQL.FieldByName("URL").AsString

    vCorpoTexto = Chr(13)+Chr(10)+Chr(13)+Chr(10)+"Bem vindo ao sistema " + vEmpresa+"."+Chr(13)+Chr(10)
    vCorpoTexto = vCorpoTexto + "Para acessar o sistema utilize o CNPJ como login e a senha informada neste email." + Chr(13) + Chr(10)

    'vUser.ChangePasswordAndMail("Senha de acesso ao sistema " + vEmpresa, vCorpoTexto)

    Dim Mail As Object
    Set Mail = NewMail
    Mail.SendTo = CurrentQuery.FieldByName("EMAILRESPONSAVEL").AsString
    Mail.Subject = "Senha de acesso ao sistema " + vEmpresa

    If CurrentQuery.FieldByName("EHFORNECEDOR").AsString = "S" Then
      Mail.Text.Add("Prezado Fornecedor" + Chr(13)+Chr(10)+Chr(13)+Chr(10))
    End If

    Mail.Text.Add(vCorpoTexto)
    Mail.Text.Add("Login: " + CurrentQuery.FieldByName("CNPJCPF").AsString)
    Mail.Text.Add("Senha: " + CurrentQuery.FieldByName("CNPJCPF").AsString + Chr(13)+Chr(10)+Chr(13)+Chr(10))
    Mail.Text.Add("Acesse: " + url)

    Mail.Send

    Set Mail = Nothing

    CriaUsuario = viHandle

    SQL.Clear
    SQL.Add("UPDATE Z_GRUPOUSUARIOS SET CODIGO = HANDLE WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = viHandle
    SQL.ExecSQL


    SQL.Clear
    SQL.Add("INSERT INTO Z_GRUPOUSUARIOS_PESSOA(HANDLE, PESSOA, USUARIO)")
    SQL.Add("VALUES (:HANDLE, :PESSOA, :USUARIO)")
    SQL.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOS_PESSOA")
    SQL.ParamByName("PESSOA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ParamByName("USUARIO").AsInteger = viHandle
    SQL.ExecSQL
  Else
    viHandle = SQL.FieldByName("HANDLE").AsInteger

    If SQL.FieldByName("PESSOA").AsInteger > 0 Then
      bsShowMessage("Este usuário já existe no sistema. Não é possível cadastrá-lo novamente", "I")
      Exit Function
    End If

    If SQL.FieldByName("MATRICULAUNICA").AsInteger > 0 Then
      Dim result As VbMsgBoxResult
      result = bsShowMessage("Este usuário está cadastrado como um usuário Beneficiário no sistema. Deseja incluí-lo também como usuário Pessoa?", "Q")
      If result = vbYes Then
        Dim SQLAUX As Object
        Set SQLAUX = NewQuery
        SQLAUX.Clear
        SQLAUX.Add("INSERT INTO Z_GRUPOUSUARIOS_PESSOA(HANDLE, PESSOA, USUARIO)")
        SQLAUX.Add("VALUES (:HANDLE, :PESSOA, :USUARIO)")
        SQLAUX.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOS_PESSOA")
        SQLAUX.ParamByName("PESSOA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQLAUX.ParamByName("USUARIO").AsInteger = viHandle
        SQLAUX.ExecSQL

        vbIncluiu = True

        SQLAUX.Clear
        SQLAUX.Add("INSERT INTO Z_GRUPOUSUARIOGRUPOS (HANDLE, GRUPO, GRUPOADICIONADO, USUARIO)")
        SQLAUX.Add("VALUES (:HANDLE, :GRUPO, :GRUPOADICIONADO, :USUARIO)")
        SQLAUX.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOGRUPOS")
        SQLAUX.ParamByName("GRUPO").AsInteger = pGrupo
        SQLAUX.ParamByName("GRUPOADICIONADO").AsInteger = vGrupo
        SQLAUX.ParamByName("USUARIO").AsInteger = viHandle
        SQLAUX.ExecSQL

        Set SQLAUX = Nothing

      End If
    End If
    If (SQL.FieldByName("PRESTADOR").AsInteger > 0) And (Not vbIncluiu) Then
      Dim vRes As VbMsgBoxResult
      vRes = bsShowMessage("Este usuário está cadastrado como um usuário Prestador no sistema. Deseja incluí-lo também como usuário Pessoa?","Q")

      If vRes = vbYes Then
        SQL.Clear
        SQL.Add("INSERT INTO Z_GRUPOUSUARIOS_PESSOA(HANDLE, PESSOA, USUARIO)")
        SQL.Add("VALUES (:HANDLE, :PESSOA, :USUARIO)")
        SQL.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOS_PESSOA")
        SQL.ParamByName("PESSOA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQL.ParamByName("USUARIO").AsInteger = viHandle
        SQL.ExecSQL

        SQL.Clear
        SQL.Add("INSERT INTO Z_GRUPOUSUARIOGRUPOS (HANDLE, GRUPO, GRUPOADICIONADO, USUARIO)")
        SQL.Add("VALUES (:HANDLE, :GRUPO, :GRUPOADICIONADO, :USUARIO)")
        SQL.ParamByName("HANDLE").AsInteger = NewHandle("Z_GRUPOUSUARIOGRUPOS")
        SQL.ParamByName("GRUPO").AsInteger = pGrupo
        SQL.ParamByName("GRUPOADICIONADO").AsInteger = vGrupo
        SQL.ParamByName("USUARIO").AsInteger = viHandle
        SQL.ExecSQL

      End If


    End If

  End If

  If (viHandle > 0) And InTransaction Then
    ' Soh fecha a transação se inseriu de fato
    If InTransaction Then
      Commit
    End If
    bsShowMessage("Usuário incluído com sucesso. Login: " + pLogin, "I")

'    If MsgBox ("Deseja enviar email ao fornecedor?",vbYesNo) = vbYes Then
'      Dim Mail As Object
'      Set Mail = NewMail
'      Mail.ShowForm("Envie o email ao prestador")
'      Mail.Send
'      Set Mail = Nothing
'    End If
' O  Cliente TJDF preferiu enviar o email automático
  End If


  Set vUser = Nothing
  Set SQL  =  Nothing
  Exit Function

  Erro:
  bsShowMessage(Err.Description, "I")
  CriaUsuario = 0

  Set vUser = Nothing
  Set SQL  =  Nothing
  If InTransaction Then
    Rollback
  End If
End Function

Public Sub BOTAOCRIARUSUARIO_OnClick()
  Dim vGrupo As String
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT GRUPOSEGURANCAPESSOA GRUPO FROM SAM_PARAMETROSWEB")
  SQL.Active = True
  vGrupo = SQL.FieldByName("GRUPO").AsString
  Set SQL = Nothing

  If CriaUsuario(CurrentQuery.FieldByName("NOME").AsString,vGrupo,CurrentQuery.FieldByName("EMAILRESPONSAVEL").AsString, CurrentQuery.FieldByName("CNPJCPF").AsString,Str(CurrentBranch)) > 0 Then
 	 Set SQL = NewQuery
     SQL.Clear
     SQL.Add("UPDATE SFN_PESSOA SET FORNECEDORCOTACAOACEITO = 'S', USUARIOFORNECEDORANALISE = :PUSUARIO, DATAFORNECEDORANALISE = :PDATAATUAL WHERE HANDLE = :HANDLE ")
     SQL.ParamByName("PUSUARIO").AsInteger = CurrentUser
	 SQL.ParamByName("PDATAATUAL").AsDateTime = ServerNow
     SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
     SQL.ExecSQL
     Set SQL = Nothing
     RefreshNodesWithTable("SFN_PESSOA")
  End If

End Sub

Public Sub BOTAOENDERECO_OnClick()
	If VisibleMode Then
		Dim viHEnderecoCpfCnpj As Long
		Dim viHEnderecoCorrespondencia As Long
		viHEnderecoCpfCnpj = CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger
		viHEnderecoCorrespondencia = CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger

		Dim dllBSInterface0028 As Object
		Set dllBSInterface0028 = CreateBennerObject("BSInterface0028.Endereco")

		Dim msgErro As String
		Dim numErro As Long

		On Error GoTo Except
			Dim vHandlePessoa As Long
			Dim vsMensagem As String

			Select Case CurrentQuery.State
				Case 1	'Browsing - não precisa editar registro da pessoa nem controlar transação
					vHandlePessoa = CurrentQuery.FieldByName("HANDLE").AsInteger

				Case 2	'Editing
					vHandlePessoa = CurrentQuery.FieldByName("HANDLE").AsInteger
					If Not InTransaction Then
						StartTransaction
					End If
				Case 3	'Inserting
					vHandlePessoa = 0
					If Not InTransaction Then
						StartTransaction
					End If
			End Select

			If vsXMLContainerEnderecos = "Vazio" Then
				vsXMLContainerEnderecos = ""
			End If
			If vsXMLEnderecosExcluidos = "Vazio" Then
				vsXMLEnderecosExcluidos = ""
			End If


			Dim vResultado As Long
			vResultado = dllBSInterface0028.Pessoa( CurrentSystem, vHandlePessoa, viHEnderecoCpfCnpj, _
				viHEnderecoCorrespondencia, vsXMLContainerEnderecos, vsXMLEnderecosExcluidos, vsMensagem)

			If vResultado = 1 Then
				vsXMLContainerEnderecos = ""
				vsXMLEnderecosExcluidos = ""
				Err.Raise(1, Err, vsMensagem)
			Else
				Select Case CurrentQuery.State
					Case 1
						'Pessoa não está em edição
						If (viHEnderecoCpfCnpj <> CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger) Or _
							(viHEnderecoCorrespondencia <> CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger) Then
							'só altera se houver troca de registro de endereço
							GravaEnderecos (CurrentQuery.FieldByName("HANDLE").AsInteger, viHEnderecoCpfCnpj, viHEnderecoCorrespondencia)
						End If
						vsXMLContainerEnderecos = ""
						vsXMLEnderecosExcluidos = ""
					Case 2, 3
						'Efetua preenchimento com os novos valores
						If viHEnderecoCpfCnpj > 0 Then
							CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger = viHEnderecoCpfCnpj
						Else
							CurrentQuery.FieldByName("ENDERECOCPFCNPJ").Clear
						End If

						If viHEnderecoCorrespondencia > 0 Then
							CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger = viHEnderecoCorrespondencia
						Else
							CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").Clear
						End If
				End Select
			End If

			AtualizaRotulosEndereco( viHEnderecoCpfCnpj, viHEnderecoCorrespondencia)
			Set dllBSInterface0028 = Nothing
			Exit Sub
		Except:
			msgErro = Err.Description
			numErro = Err.Number
			Set dllBSInterface0028 = Nothing
			UpdateLastUpdate("SFN_PESSOA")
			AtualizaRotulosEndereco( CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger, CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger )
			bsShowMessage("Falha no cadastro de Endereços: "+Chr(13) +"(" + CStr( numErro) +")"+ msgErro, "E")
	End If
End Sub

Public Sub BOTAOFINANCEIRO_OnClick()
  Dim SQL As Object
  Dim Interface As Object

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SFN_CONTAFIN WHERE PESSOA=" + CurrentQuery.FieldByName("HANDLE").AsString)
  SQL.Active = True
  If Not SQL.EOF Then
    Set Interface = CreateBennerObject("SamContaFinanceira.Consulta")
    Interface.Exec(CurrentSystem, SQL.FieldByName("HANDLE").AsInteger)
    Set Interface = Nothing
  Else
    bsShowMessage("Conta financeira não encontrada", "I")
  End If
  SQL.Active = False

  Set SQL = Nothing

End Sub

Public Sub EHGRUPOFATURAMENTO_OnChange()
  If CurrentQuery.FieldByName("EHGRUPOFATURAMENTO").AsString = "S" Then
    CurrentQuery.FieldByName("GRUPOFATURAMENTO").Clear
  End If
End Sub

Public Sub TABFISICAJURIDICA_OnChange()
  If CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger = 1 Then
    CurrentQuery.FieldByName("GRUPOFATURAMENTO").Clear
  End If
End Sub

Public Sub TABFISICAJURIDICA_OnChanging(AllowChange As Boolean)
  CurrentQuery.FieldByName("CNPJCPF").Clear
  If TABFISICAJURIDICA.PageIndex = 1 Then
    CurrentQuery.FieldByName("CNPJCPF").Mask = "999\.999\.999\-99;0;_"
  Else
    CurrentQuery.FieldByName("CNPJCPF").Mask = "99\.999\.999\/9999\-99;0;_"
  End If
End Sub

Public Sub TABLE_AfterCancel()
	If InTransaction Then
		Rollback
	End If
	vsXMLContainerEnderecos = ""
	vsXMLEnderecosExcluidos = ""
End Sub

Public Sub TABLE_AfterCommitted()
  ' VERIFICA A CONTA FINANCEIRA
  Dim Erro As Long
  Dim InterfaceFin As Object

  If Not InTransaction Then
    StartTransaction
  End If
  Set InterfaceFin = CreateBennerObject("FINANCEIRO.ContaFin")
  Erro = InterfaceFin.Cadastro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 3, 0)
  If Erro <= 0 Then
    bsShowMessage("Erro " + Str(Erro) + " ao criar Conta Financeira", "I")
  End If
  Set InterfaceFin = Nothing

  If InTransaction Then
    Commit
  End If
  'FIM VERIFICA A CONTA FINANCEIRA

  AtualizaRotulosEndereco( CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger, CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger )

End Sub


Public Sub TABLE_AfterDelete()
  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SFN_PESSOA")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "Z")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterInsert()
  If (Not (WebMode)) And (VisibleMode) Then
    CurrentQuery.FieldByName("CNPJCPF").Mask = "999\.999\.999\-99;0;_"
  End If

  'Exclusivo web
  If WebVisionCode = "WEBFORNECEDORINC" Then
    CurrentQuery.FieldByName("EHFORNECEDOR").AsString = "S"
    CurrentQuery.FieldByName("FORNECEDORWEB").AsString = "S"
    CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger = 2

    Dim sqlx As Object
    Set sqlx = NewQuery
    sqlx.Clear
    sqlx.Add("SELECT LIVREESCOLHAISSJURIDICA FROM SAM_PARAMETROSPRESTADOR")
    sqlx.Active = True

    CurrentQuery.FieldByName("ISS").AsInteger = sqlx.FieldByName("LIVREESCOLHAISSJURIDICA").AsInteger

    sqlx.Clear
    sqlx.Add("SELECT MIN(HANDLE) HANDLE FROM SAM_ENDERECO")
    sqlx.Active = True

    CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger = sqlx.FieldByName("HANDLE").AsInteger

    Set sqlx = Nothing

    If (WebMode) Then
      CurrentQuery.FieldByName("TELEFONE1").Mask = ""
    End If

  End If
End Sub

Public Sub TABLE_AfterPost()
  AtualizaRotulosEndereco( CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger, CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger )

  Dim AtualizaPessoa As Object
  Dim vsMensagem     As String
  Set AtualizaPessoa = CreateBennerObject("SamBeneficiario.Atualiza")
  AtualizaPessoa.Pessoa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,0, vsMensagem)
  If vsMensagem <> "" Then
    bsShowMessage("Atualização da Região não finalizada : " + vsMensagem, "I")
  End If
  Set AtualizaPessoa = Nothing

  If vsModoEdicao = "A" Then
    Dim viRetorno   As Integer
 	Dim dllBSBen021 As Object
	Set dllBSBen021 = CreateBennerObject("BSBen021.AtualizacaoEndereco")

	viRetorno = dllBSBen021.Excluir(CurrentSystem, _
	                                vsXMLEnderecosExcluidos, _
	                                vsMensagem)

	Set dllBSBen021 = Nothing

	If viRetorno = 1 Then
      Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Gravação cancelada!")
	Else
	  If vsMensagem <> "" Then
	    bsShowMessage(vsMensagem, "I")
	  End If
	End If
  End If

  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SFN_PESSOA")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterScroll()
	BOTAOENDERECO.Visible     = True
  If CurrentQuery.State <> 1 Then
    BOTAOFINANCEIRO.Visible   = False
    BOTAOCRIARUSUARIO.Visible = False
    'BOTAOENDERECO.Visible     = True
  Else
    BOTAOFINANCEIRO.Visible   = True
    BOTAOCRIARUSUARIO.Visible = True
    'BOTAOENDERECO.Visible     = False
  End If


  Dim CNPJCPF As String

  CNPJCPF = CurrentQuery.FieldByName("CNPJCPF").AsString
  If Len(CNPJCPF) = 11 Then
    CurrentQuery.FieldByName("CNPJCPF").Mask = "999\.999\.999\-99;0;_"
  ElseIf Len(CNPJCPF) = 14 Then
    CurrentQuery.FieldByName("CNPJCPF").Mask = "99\.999\.999\/9999\-99;0;_"
  End If

  AtualizaRotulosEndereco( CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger, CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger )
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vsModoEdicao = "A"
  vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""

  BOTAOFINANCEIRO.Visible   = False
  BOTAOCRIARUSUARIO.Visible = False
  'BOTAOENDERECO.Visible     = True

  Dim CNPJCPF As String

  CNPJCPF = CurrentQuery.FieldByName("CNPJCPF").AsString
  If Len(CNPJCPF) = 11 Then
    CurrentQuery.FieldByName("CNPJCPF").Mask = "999\.999\.999\-99;0;_"
  ElseIf Len(CNPJCPF) = 14 Then
    CurrentQuery.FieldByName("CNPJCPF").Mask = "99\.999\.999\/9999\-99;0;_"
  End If

  If (WebMode) Then
      CurrentQuery.FieldByName("TELEFONE1").Mask = ""
  End If

  'Se estiver em modo desktop a transação deve ser iniciada antes da edição
  'pela possibilidade de inclusão/alteração de endereços
  If VisibleMode And Not InTransaction Then
    StartTransaction
  End If
End Sub

'#Uses "*VerificaEmail"

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
    'Em modo web o tratamento de máscara será por JavaScript
    CurrentQuery.FieldByName("CNPJCPF").Mask = ""
  Else
    'Como na inclusão assume-se como Física a máscara inicial será de CPF
    If CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger = 1 Then
      CurrentQuery.FieldByName("CNPJCPF").Mask = "999\.999\.999\-99;0;_"
    Else
      CurrentQuery.FieldByName("CNPJCPF").Mask = "99\.999\.999\/9999\-99;0;_"
    End If
  End If

  vsModoEdicao = "I"
  vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""

  BOTAOFINANCEIRO.Visible   = False
  BOTAOCRIARUSUARIO.Visible = False
  'BOTAOENDERECO.Visible     = True

  'Se estiver em modo desktop a transação deve ser iniciada antes da edição
  'pela possibilidade de inclusão/alteração de endereços
  If VisibleMode And Not InTransaction Then
    StartTransaction
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Remover a formação do campo de CPF/CNPJ na inclusão de registros em modo Web
  If WebMode And _
     CurrentQuery.State Then
    CurrentQuery.FieldByName("CNPJCPF").AsString = Replace(Replace(Replace(CurrentQuery.FieldByName("CNPJCPF").AsString, ".", ""), "/", ""), "-", "")
  End If

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  If (CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger = 2) And _
     Not CurrentQuery.FieldByName("ENDERECOCPFCNPJ").IsNull Then
    If Not CurrentQuery.FieldByName("INSCRICAOESTADUAL").IsNull Then
      Dim obj As Object
      Dim Sigla As String

      Set obj = CreateBennerObject("SAMUTIL.ROTINAS")

      Dim sqlx As Object
      Set sqlx = NewQuery

      sqlx.Clear
      sqlx.Add("SELECT SIGLA FROM ESTADOS WHERE HANDLE = (")
      sqlx.Add("SELECT ESTADO FROM SAM_ENDERECO WHERE HANDLE = :HANDLE)")
      sqlx.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger
      sqlx.Active = True
      Sigla  = sqlx.FieldByName("SIGLA").AsString

      If Not obj.ValidarInscricaoEstadual(CurrentSystem, Sigla, CurrentQuery.FieldByName("INSCRICAOESTADUAL").AsString) Then
        If (WebMode) And (WebMenuCode <> "") Then
    	  CancelDescription = "Inscrição estadual informada inválida"
        Else
          bsShowMessage("Inscrição estadual informada inválida", "E")
        End If
        CanContinue = False
        Exit Sub
      End If
    End If
  End If

  If WebMode Then
    Dim Telefone As String
    Dim VFAX As String

	If Not CurrentQuery.FieldByName("TELEFONE1").IsNull Then
      Telefone = FormatarTelefone(CurrentQuery.FieldByName("DDD1").AsString,CurrentQuery.FieldByName("TELEFONE1").AsString)
      If Mid(Telefone,1,1) <> "(" Then
      	CanContinue = False
      	CancelDescription = Telefone
      	CurrentQuery.FieldByName("TELEFONE2").AsString = ""
      Else
      	CurrentQuery.FieldByName("TELEFONE1").AsString = Telefone
      	CurrentQuery.FieldByName("TELEFONE2").AsString = ""
      	'CurrentQuery.FieldByName("DDD1").AsString = ""
      End If
    End If

    If Not CurrentQuery.FieldByName("FAX").IsNull Then
      VFAX  = FormatarTelefone(CurrentQuery.FieldByName("PREFIXO1").AsString,CurrentQuery.FieldByName("FAX").AsString)
      If Mid(VFAX,1,1) <> "(" Then
        CanContinue = False
        CancelDescription = VFAX
        CurrentQuery.FieldByName("FAX").AsString = ""
      Else
        CurrentQuery.FieldByName("FAX").AsString = VFAX
      End If
    End If
  End If

  'Somente para web
  If WebVisionCode = "WEBFORNECEDORINC" Then
    If Len(CurrentQuery.FieldByName("CNPJCPF").AsString) = 11 Then
      CancelDescription = "Somente pessoas jurídicas podem participar de cotação de preços"
      CanContinue = False
    Else

      If CurrentQuery.FieldByName("CNPJCPF").AsString <> "" Then
	      If Not IsValidCGC(CurrentQuery.FieldByName("CNPJCPF").AsString) Then
	        CancelDescription = "CNPJ inválido"
	        CanContinue = False
	      End If
	  End If
    End If

    Dim SQLX1 As Object
    Set SQLX1 = NewQuery
    SQLX1.Add("SELECT NOME FROM SFN_PESSOA WHERE CNPJCPF = :CNPJCPF AND HANDLE <> :HANDLE ")
    SQLX1.ParamByName("CNPJCPF").Value = CurrentQuery.FieldByName("CNPJCPF").AsString
    SQLX1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQLX1.Active = True
    If Not SQLX1.EOF Then
      CanContinue = False
      CancelDescription = "A pessoa '" + SQLX1.FieldByName("NOME").AsString + "' já possui esse CPF/CNPJ"
      SQLX1.Active = False
      Set SQLX1 = Nothing
      Exit Sub
    End If

    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("EMAIL").IsNull Then
    If Not VerificaEmail(CurrentQuery.FieldByName("EMAIL").AsString)Then
      bsShowMessage("E-mail inválido", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If Not CurrentQuery.FieldByName("EMAILRESPONSAVEL").IsNull Then
    If Not VerificaEmail(CurrentQuery.FieldByName("EMAILRESPONSAVEL").AsString)Then
      bsShowMessage("E-mail responsável inválido", "E")
      CanContinue = False
      Exit Sub
    End If
  End If


  If Not CurrentQuery.FieldByName("CNPJCPF").IsNull Then
    If CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger = 1 Then
      If Not IsValidCPF(CurrentQuery.FieldByName("CNPJCPF").AsString) Then
        CanContinue = False
        bsShowMessage("CPF inválido", "E")
        Exit Sub
      End If
    Else
      If Not IsValidCGC(CurrentQuery.FieldByName("CNPJCPF").AsString) Then
        CanContinue = False
        bsShowMessage("CNPJ inválido", "E")
        Exit Sub
      End If
    End If
  End If

  Dim Pres As Object
  Set Pres = NewQuery
  Pres.Add("SELECT NOME FROM SFN_PESSOA WHERE CNPJCPF = :CNPJCPF AND HANDLE <> :HANDLE ")
  Pres.ParamByName("CNPJCPF").Value = CurrentQuery.FieldByName("CNPJCPF").AsString
  Pres.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  Pres.Active = True
  If Not Pres.EOF Then
    CanContinue = False
    bsShowMessage("A pessoa " + Pres.FieldByName("NOME").AsString + " já possui esse CPF/CNPJ", "E")
    Pres.Active = False
    Set Pres = Nothing
    Exit Sub
  End If

  TABLE_AfterScroll

  If((CurrentQuery.FieldByName("EHCONTRATO").AsString = "S")Or _
     (CurrentQuery.FieldByName("EHCONVENIO").AsString = "S")Or _
     (CurrentQuery.FieldByName("EHFISCO").AsString = "S"))And _
     (CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger <>2)Then
    bsShowMessage("Contrato, Convênio e Fisco devem ser do tipo jurídica", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim especifico As Object
  Set especifico = CreateBennerObject("ESPECIFICO.uEspecifico")

  If(CurrentQuery.FieldByName("EHRESPONSAVELPORFAMILIA").AsString = "S")And _
     ( Not especifico.BEN_VerificaTipoPessoaResponsavelFamilia(CurrentSystem, CurrentQuery.FieldByName("TABFISICAJURIDICA").AsString)) Then
    bsShowMessage("Responsável por família deve ser do tipo física", "E")
    CanContinue = False
    Set especifico = Nothing
    Exit Sub
  End If
  Set especifico = Nothing

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT FISICAJURIDICA FROM SFN_ISS WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ISS").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    If SQL.FieldByName("FISICAJURIDICA").AsInteger = 1 And _
       CurrentQuery.FieldByName("TABFISICAJURIDICA").Value <>1 Then
      Set SQL = Nothing
      bsShowMessage("Para o tipo de ISS informado a pessoa deve ser física", "E")
      CanContinue = False
      Exit Sub
    End If
    If SQL.FieldByName("FISICAJURIDICA").AsInteger = 2 And _
       CurrentQuery.FieldByName("TABFISICAJURIDICA").Value <>2 Then
      Set SQL = Nothing
      bsShowMessage("Para o tipo de ISS informado a pessoa deve ser jurídica", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  Set SQL = Nothing

  If CurrentQuery.FieldByName("EHGRUPOFATURAMENTO").AsString = "N" Then
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SFN_PESSOA")
    SQL.Add("WHERE GRUPOFATURAMENTO = :HPESSOA")
    SQL.ParamByName("HPESSOA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If Not SQL.EOF Then
      CanContinue = False
      Set SQL = Nothing
      bsShowMessage("Esta pessoa tem ligação de grupo de faturamento em outras pessoas. O indicador 'Grupo faturamento' não pode ser desmarcado", "E")
      Exit Sub
    End If
  Else
    If CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger = 1 Then
      CanContinue = False
      bsShowMessage("Para marcar o indicador 'Grupo faturamento' a pessoa não pode ser 'Física'", "E")
      Exit Sub
    End If

    If Not(CurrentQuery.FieldByName("GRUPOFATURAMENTO").IsNull)And _
       CurrentQuery.FieldByName("TABFISICAJURIDICA").AsInteger = 1 Then
      CanContinue = False
      bsShowMessage("Não pode referenciar uma pessoa para grupo de faturamento se for do tipo 'Física'", "E")
      Exit Sub
    End If

    If Not(CurrentQuery.FieldByName("GRUPOFATURAMENTO").IsNull)Then
      CanContinue = False
      bsShowMessage("Não pode referenciar uma pessoa para grupo de faturamento se ele própria estiver marcada como grupo de faturamento", "E")
      Exit Sub
    End If
  End If

  If (Not CurrentQuery.FieldByName("DATASAIDA").IsNull)And _
     (CurrentQuery.FieldByName("DATASAIDA").AsDateTime <CurrentQuery.FieldByName("DATAENTRADA").AsDateTime)Then
    bsShowMessage("A Data Saída , se informada, deve ser maior ou igual a entrada", "E")
    CanContinue = False
  Else
    CanContinue = True
  End If

'lopes - sms 54951
  If Not CurrentQuery.FieldByName("ENDERECOCPFCNPJ").IsNull Then

     Dim qSel As Object
     Set qSel = NewQuery
     qSel.Clear
     qSel.Add(" SELECT M.REGIAO              ")
     qSel.Add("   FROM SAM_ENDERECO E,       ")
     qSel.Add("        MUNICIPIOS   M        ")
     qSel.Add("  WHERE E.HANDLE = :ENDERECO  ")
     qSel.Add("    AND M.HANDLE = E.MUNICIPIO")
     qSel.ParamByName("ENDERECO").AsInteger = CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsInteger
     qSel.Active =True

     If qSel.FieldByName("REGIAO").IsNull Then
       CanContinue = False
       bsShowMessage("O endereço informado não tem um município válido, com região e estado cadastrados!", "E")
       Set qSel = Nothing
       Exit Sub
     End If

    Set qSel = Nothing
  Else
    'Verificação de obrigatoriedade do endereço de CNPJ/CPF
    If (WebMode And CurrentQuery.State = 3) Then
      bsShowMessage("Pessoa exige informações de endereço do CNPJ/CPF!", "I")
    Else
      CanContinue = False
      bsShowMessage("Pessoa exige informações de endereço do CNPJ/CPF!", "E")
    End If
  End If
End Sub

Public Function GravaEnderecos(pPessoa As Long, pEnderecoCnpjCpf As Long, pEnderecoCorrespondencia As Long)
	Dim iniciouTransacao As Boolean
	Dim msgErro As String
	Dim numErro As Long

	On Error GoTo Except
		Dim sqlUp As Object
		Set sqlUp = NewQuery

		If Not pPessoa > 0 Then
			Err.Raise(1, Err, "Falha ao atualizar endereços: Pessoa não informada!")
		End If
		If Not (pEnderecoCnpjCpf > 0) And Not (pEnderecoCorrespondencia > 0) Then
			If CurrentQuery.FieldByName("EHRESPONSAVELPORFAMILIA").AsString = "S" Then
				Err.Raise(1, Err, "Para responsáveis por família um dos endereços é obrigatório!")
			End If
		End If
		Dim vEndPes As String
		Dim vEndCor As String
		If pEnderecoCnpjCpf > 0 Then
			vEndPes = CStr(pEnderecoCnpjCpf)
		Else
			vEndPes = "NULL"
		End If
		If pEnderecoCorrespondencia > 0 Then
			vEndCor = CStr(pEnderecoCorrespondencia)
		Else
			vEndCor = "NULL"
		End If

		sqlUp.Add("UPDATE SFN_PESSOA SET ENDERECOCPFCNPJ = " + vEndPes + ", ENDERECOCORRESPONDENCIA = " + vEndCor + " WHERE HANDLE = :HANDLE ")
		sqlUp.ParamByName("HANDLE").AsInteger = pPessoa

		iniciouTransacao = False
		If Not InTransaction Then
			StartTransaction
			iniciouTransacao = True
		End If

		sqlUp.ExecSQL

		If iniciouTransacao And InTransaction Then
			Commit
			iniciouTransacao = False
		End If

		Set sqlUp = Nothing
		Exit Function
	Except:
		msgErro = Err.Description
		numErro = Err.Number
		Set sqlUp = Nothing
		If iniciouTransacao And InTransaction Then
			Rollback
		End If
		Err.Raise(numErro, Err, msgErro)
End Function

Public Function preencheValor(pPrefixo As String, pValor As String, pSufixo As String, pValorSeVazio) As String
	If pValor <> "" Then
		preencheValor = pPrefixo + pValor + pSufixo
	Else
		preencheValor = pValorSeVazio
	End If
End Function

Public Function preencheRotulosEndereco(pTipo As String, pRot0 As String, pRot1 As String, pRot2 As String, pRot3 As String, pRot4 As String, pRot5 As String)
	Select Case pTipo
		Case "CPF"
	    	ROTULOCNPJCPF1.Text = pRot0 + pRot1
	    	ROTULOCNPJCPF2.Text = pRot2
	    	ROTULOCNPJCPF3.Text = pRot3
	    	ROTULOCNPJCPF4.Text = pRot4
	    	ROTULOCNPJCPF5.Text = pRot5
		Case "COR"
		  	ROTULOCORRESP1.Text = pRot0 + pRot1
		  	ROTULOCORRESP2.Text = pRot2
			ROTULOCORRESP3.Text = pRot3
		 	ROTULOCORRESP4.Text = pRot4
		 	ROTULOCORRESP5.Text = pRot5
	End Select
End Function

Public Function AtualizaRotulosEndereco( pEnderecoCnpjCpf As Long, pEnderecoCorrespondencia As Long)
	Dim vQryEndereco As Object
	Dim vListaEnderecos As String
	vListaEnderecos = ""

	If pEnderecoCnpjCpf > 0 Then
		vListaEnderecos  = CStr( pEnderecoCnpjCpf )
	Else
		preencheRotulosEndereco("CPF", "", "", "", "", "", "")
	End If

	If pEnderecoCorrespondencia > 0 Then
		If vListaEnderecos <> "" Then
			vListaEnderecos = vListaEnderecos + ", " + CStr( pEnderecoCorrespondencia )
		Else
			vListaEnderecos = CStr( pEnderecoCorrespondencia )
		End If
	Else
		preencheRotulosEndereco("COR", "", "", "", "", "", "")
	End If

	If (vListaEnderecos <> "") Then
		On Error GoTo Except
			Dim vLogradouro, vNumero, vComplemento, vBairro, vCEP, vTelefone1, vTelefone2, vFax, vCelular, vRamal As String
			Dim vMunicipio, vEstado, vTipoLogradouro As String

			Set vQryEndereco = NewQuery
			vQryEndereco.Active = False
			vQryEndereco.Clear
			vQryEndereco.Add("SELECT E.HANDLE, E.ESTADO, E.MUNICIPIO, E.BAIRRO, E.CEP, E.NUMERO, E.COMPLEMENTO, E.TELEFONE1, ")
			vQryEndereco.Add("       E.TELEFONE2, E.FAX, E.LOGRADOURO, E.CELULAR, E.RAMAL, LT.DESCRICAO AS TIPOLOGRADOURO,   ")
			vQryEndereco.Add("       ES.NOME NOMEESTADO, M.NOME NOMEMUNICIPIO ")
			vQryEndereco.Add("  FROM SAM_ENDERECO E")
			vQryEndereco.Add("  LEFT JOIN LOGRADOUROS_TIPO LT ON LT.HANDLE = E.TIPOLOGRADOURO ")
			vQryEndereco.Add("  LEFT JOIN ESTADOS ES ON ES.HANDLE = E.ESTADO ")
			vQryEndereco.Add("  LEFT JOIN MUNICIPIOS M ON M.HANDLE = E.MUNICIPIO ")
			vQryEndereco.Add(" WHERE E.HANDLE IN (" + vListaEnderecos +") ")

			vQryEndereco.Active = True

			While Not vQryEndereco.EOF
				vLogradouro =	preencheValor(""			 , vQryEndereco.FieldByName("LOGRADOURO").AsString, 		""		,"")
				vNumero     =	preencheValor(", Nº "		 , vQryEndereco.FieldByName("NUMERO").AsString, 			""		,"")
				vComplemento=	preencheValor("Complemento: ",vQryEndereco.FieldByName("COMPLEMENTO").AsString, 		"     "	,"")
				vBairro 	=	preencheValor("Bairro: "	 , vQryEndereco.FieldByName("BAIRRO").AsString, 			""		,"")
				vCEP 		=	preencheValor("CEP: "		 , vQryEndereco.FieldByName("CEP").AsString, 				"     "	,"")
				vMunicipio	=	preencheValor("Município: "	 , vQryEndereco.FieldByName("NOMEMUNICIPIO").AsString,		"     "	,"")
				vEstado		=	preencheValor("Estado: "	 , vQryEndereco.FieldByName("NOMEESTADO").AsString, 		""		,"")
				vTelefone1	=	preencheValor("Telefone 1: " , vQryEndereco.FieldByName("TELEFONE1").AsString, 			"     "	,"")
				vTelefone2	=	preencheValor("Telefone 2: " , vQryEndereco.FieldByName("TELEFONE2").AsString, 			"     "	,"")
				vRamal		=	preencheValor("Ramal: "		 , vQryEndereco.FieldByName("RAMAL").AsString, 				"     "	,"")
				vFax		=	preencheValor("Fax: "		 , vQryEndereco.FieldByName("FAX").AsString, 				""		,"")
				vCelular	=	preencheValor("Celular: "	 , vQryEndereco.FieldByName("CELULAR").AsString, 			"     "	,"")
				vTipoLogradouro = preencheValor(""			 , vQryEndereco.FieldByName("TIPOLOGRADOURO").AsString,		" "		,"")

				Select Case vQryEndereco.FieldByName("HANDLE").AsInteger
					Case pEnderecoCnpjCpf
						preencheRotulosEndereco("CPF", vTipoLogradouro + vLogradouro, _
													   vNumero, _
													   vComplemento + vBairro, _
													   vCEP + vMunicipio + vEstado, _
													   vTelefone1 + vTelefone2 + vRamal + vFax, _
													   vCelular)
						If pEnderecoCorrespondencia = pEnderecoCnpjCpf Then
							GoTo copiaParaCorrespondencia
						End If

					Case pEnderecoCorrespondencia
						copiaParaCorrespondencia:
						preencheRotulosEndereco("COR", vTipoLogradouro + vLogradouro, _
													   vNumero, _
													   vComplemento + vBairro, _
													   vCEP + vMunicipio + vEstado, _
													   vTelefone1 + vTelefone2 + vRamal + vFax, _
													   IIf( vCelular <> "", vCelular + "     ", ""))
				End Select

				vQryEndereco.Next
			Wend
			vQryEndereco.Active = False
			Set vQryEndereco = Nothing

			Exit Function

		Except:
			Set vQryEndereco = Nothing
			Err.Raise(Err.Number, Err.Source, "Falha ao exibir endereços do Beneficiário: " + Err.Description)
	End If
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCRIARUSUARIO"
			BOTAOCRIARUSUARIO_OnClick
		Case "BOTAOENDERECO"
			BOTAOENDERECO_OnClick
		Case "BOTAOFINANCEIRO"
			BOTAOFINANCEIRO_OnClick
	End Select
End Sub
