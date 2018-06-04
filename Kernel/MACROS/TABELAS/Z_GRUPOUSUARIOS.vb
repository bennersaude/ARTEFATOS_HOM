'HASH: 196516B6F898E565A9927B65661E60E4
'Macro: Z_GRUPOUSUARIOS
'#Uses "*bsShowMessage"

Dim iGrupo       As Integer
Dim vsModoEdicao As String

Public Sub BOTAOLIBERARGLOSA_OnClick()
  Dim SQL1 As Object
  Dim SQL2 As Object

  If bsShowMessage("Confirma a liberação de todas as glosas para o usuário ?", "Q") = vbYes Then
	Set SQL1 = NewQuery
	Set SQL2 = NewQuery


	SQL1.Clear
	SQL1.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA")
	SQL1.Add("WHERE HANDLE NOT IN")
	SQL1.Add("(SELECT MOTIVOGLOSA FROM SAM_USUARIO_MOTIVOGLOSA WHERE USUARIO = :USUARIO)")
	SQL1.ParamByName("USUARIO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger  'RecordHandleOfTable("Z_GRUPOUSUARIOS")
	SQL1.Active=True

	While Not SQL1.EOF
	  SQL2.Clear
	  SQL2.Add("INSERT INTO SAM_USUARIO_MOTIVOGLOSA (HANDLE, USUARIO, MOTIVOGLOSA) VALUES (:HANDLE,:USUARIO,:MOTIVOGLOSA)")
	  SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVOGLOSA")
	  SQL2.ParamByName("USUARIO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger  'RecordHandleOfTable("Z_GRUPOUSUARIOS")
	  SQL2.ParamByName("MOTIVOGLOSA").Value = SQL1.FieldByName("HANDLE").AsInteger
	  SQL2.ExecSQL
	  SQL1.Next
	Wend

    Set SQL1 = Nothing
    Set SQL2 = Nothing

  End If

End Sub

Public Sub BOTAOLIBERARNEGACAO_OnClick()
  Dim SQL1 As Object
  Dim SQL2 As Object

  If bsShowMessage("Confirma a liberação de todas as negações para o usuário ?" , "Q") = vbYes Then
    Set SQL1 = NewQuery
	Set SQL2 = NewQuery


	SQL1.Clear
	SQL1.Add("SELECT HANDLE FROM SAM_MOTIVONEGACAO")
	SQL1.Add("WHERE HANDLE NOT IN")
	SQL1.Add("(SELECT MOTIVONEGACAO FROM SAM_USUARIO_MOTIVONEGACAO WHERE USUARIO = :USUARIO)")
	SQL1.ParamByName("USUARIO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger  'RecordHandleOfTable("Z_GRUPOUSUARIOS")
	SQL1.Active=True

	While Not SQL1.EOF
	  SQL2.Clear
	  SQL2.Add("INSERT INTO SAM_USUARIO_MOTIVONEGACAO (HANDLE, USUARIO, MOTIVONEGACAO) VALUES (:HANDLE,:USUARIO,:MOTIVONEGACAO)")
	  SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVONEGACAO")
	  SQL2.ParamByName("USUARIO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger  'RecordHandleOfTable("Z_GRUPOUSUARIOS")
	  SQL2.ParamByName("MOTIVONEGACAO").Value = SQL1.FieldByName("HANDLE").AsInteger
	  SQL2.ExecSQL
	  SQL1.Next
	Wend

    Set SQL1 = Nothing
    Set SQL2 = Nothing

  End If
End Sub

Public Sub BOTAONOVASENHA_OnClick()
  On Error GoTo erro

    StartTransaction

    Dim Usuario  As CSSystemUser
    Set Usuario = NewSystemUser

    Usuario.SelectUser(CurrentQuery.FieldByName("HANDLE").AsInteger)
    Usuario.ChangePasswordAndMail("Sistema Benner Saúde: Alteração da senha de acesso ao sistema", "Solicitação de nova senha para o sistema Benner. </br></br> Usuário: [USUARIO] </br> Nova senha: [SENHA]")
    Set Usuario = Nothing
    bsShowMessage("Nova senha enviada!", "I")

    Commit

    Exit Sub

  erro:
    If InTransaction Then
      Rollback
    End If
    bsShowMessage("Erro ao enviar nova senha: " + Err.Description, "I")
End Sub

Public Sub BSBOTAOCOPIAR_OnClick()
  If bsShowMessage("Deseja copiar este Usuário ? ", "Q") = vbYes Then

  	On Error GoTo FIM

  	If Not InTransaction Then
	    StartTransaction
  	End If

  	Dim HandleGrupoUsuario As Long
  	Dim HandleGrupoUsuarioEmpresas As Long
  	Dim HandleGrupoUsuarioGrupos As Long
  	Dim HandleAtual As Long
  	Dim lPassou As Boolean


  	Dim SQL As Object
  	Dim SQL1 As Object
  	Dim SQL2 As Object

  	Dim Str1 As String

  	Set SQL = NewQuery
  	Set SQL1 = NewQuery
  	Set SQL2 = NewQuery

  	lPassou = False

  	'CRIA USUµRIO
  	HandleGrupoUsuario = NewHandle("Z_GRUPOUSUARIOS")



  	'Incluídos os campos FNFATURAAVULSA ,SFNFATURABAIXAR,SFNFATURACANCELAR, SFNDOCUMENTOGERAR, SFNDOCUMENTOBAIXAR, SFNDOCUMENTOCANCELAR,
  	'             SFNDOCUMENTOIMPRIMIR, SFNPARCELAMENTO, FILIALPADRAO, IDIOMA, ALTERAR, INCLUIR, EXCLUIR, DESENVOLVER, PERMITESOBREPORHORARIO
  	'SMS 9266
  	SQL.Add("INSERT INTO Z_GRUPOUSUARIOS (HANDLE, GRUPO, APELIDO, NOME, PROTEGERREGISTRO, CODIGO, SFNFATURAAVULSA ,SFNFATURABAIXAR,	")
  	SQL.Add("  SFNFATURACANCELAR, SFNDOCUMENTOGERAR, SFNDOCUMENTOBAIXAR, SFNDOCUMENTOCANCELAR,			")
  	SQL.Add("  SFNDOCUMENTOIMPRIMIR, SFNPARCELAMENTO, ALTERAR, INCLUIR, EXCLUIR, ALTERARSENHA, INATIVO, DESENVOLVER, RETERSENHA, USUARIOREMOTO, PERMITESOBREPORHORARIO, EMAIL, ESPECIALIDADES,")

  	SQL.Add(IIf(Not CurrentQuery.FieldByName("IDIOMA").IsNull, "IDIOMA, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("TIPOAUTORIZACAOPADRAO").IsNull, "TIPOAUTORIZACAOPADRAO, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("FILIALPADRAO").IsNull, "FILIALPADRAO, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("ULTIMAEMPRESA").IsNull, "ULTIMAEMPRESA, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("ULTIMAFILIAL").IsNull, "ULTIMAFILIAL, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("ULTIMOMODULO").IsNull, "ULTIMOMODULO, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("FORMWIDTH").IsNull, "FORMWIDTH, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("FORMHEIGHT").IsNull, "FORMHEIGHT, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("MUDANCAFASE").IsNull, "MUDANCAFASE, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("SMTPPORT").IsNull, "SMTPPORT, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("SMTPSERVER").IsNull, "SMTPSERVER, ", ""))
  	'SQL.Add(IIf(Not CurrentQuery.FieldByName("MEMO").IsNull, "MEMO, ", ""))

  	SQL.Add("MEMO, UNIDADEMEDIDA, VIEWGRID, SMTPAUTENTICADO, GERENCIATODOSAGENDAMENTOS, DESENVOLVERRELATORIO) ")
  	SQL.Add("VALUES ")
  	SQL.Add("(:HANDLE, :GRUPO, :APELIDO, :NOME, :PROTEGERREGISTRO, :HANDLE, :SFNFATURAAVULSA ,:SFNFATURABAIXAR,")
  	SQL.Add("                             :SFNFATURACANCELAR, :SFNDOCUMENTOGERAR, :SFNDOCUMENTOBAIXAR, :SFNDOCUMENTOCANCELAR,			")
  	SQL.Add("                             :SFNDOCUMENTOIMPRIMIR, :SFNPARCELAMENTO, :ALTERAR, :INCLUIR,")
  	SQL.Add(" 							  :EXCLUIR, :ALTERARSENHA, :INATIVO, :DESENVOLVER, :USUARIOREMOTO, :RETERSENHA, :PERMITESOBREPORHORARIO, :EMAIL, :ESPECIALIDADES,")

  	SQL.Add(IIf(Not CurrentQuery.FieldByName("IDIOMA").IsNull, ":IDIOMA, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("TIPOAUTORIZACAOPADRAO").IsNull, ":TIPOAUTORIZACAOPADRAO, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("FILIALPADRAO").IsNull, ":FILIALPADRAO, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("ULTIMAEMPRESA").IsNull, ":ULTIMAEMPRESA, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("ULTIMAFILIAL").IsNull, ":ULTIMAFILIAL, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("ULTIMOMODULO").IsNull, ":ULTIMOMODULO, ", ""))
	SQL.Add(IIf(Not CurrentQuery.FieldByName("FORMWIDTH").IsNull, ":FORMWIDTH, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("FORMHEIGHT").IsNull, ":FORMHEIGHT, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("MUDANCAFASE").IsNull, ":MUDANCAFASE, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("SMTPPORT").IsNull, ":SMTPPORT, ", ""))
  	SQL.Add(IIf(Not CurrentQuery.FieldByName("SMTPSERVER").IsNull, ":SMTPSERVER, ", ""))

  	'SQL.Add(IIf(Not CurrentQuery.FieldByName("MEMO").IsNull, ":MEMO, ", ""))

  	SQL.Add(":MEMO, :UNIDADEMEDIDA, :VIEWGRID, :SMTPAUTENTICADO, :GERENCIATODOSAGENDAMENTOS, :DESENVOLVERRELATORIO)")

  	SQL.ParamByName("Handle").Value = HandleGrupoUsuario
  	SQL.ParamByName("SMTPAUTENTICADO").Value = CurrentQuery.FieldByName("SMTPAUTENTICADO").Value
  	SQL.ParamByName("GERENCIATODOSAGENDAMENTOS").Value = CurrentQuery.FieldByName("GERENCIATODOSAGENDAMENTOS").Value
  	SQL.ParamByName("DESENVOLVERRELATORIO").Value = CurrentQuery.FieldByName("DESENVOLVERRELATORIO").Value
  	SQL.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("GRUPO").AsInteger
  	SQL.ParamByName("APELIDO").Value = Mid(CurrentQuery.FieldByName("APELIDO").AsString + "_" + Str(HandleGrupoUsuario),1,190)
  	SQL.ParamByName("NOME").Value = Mid(CurrentQuery.FieldByName("NOME").AsString + "_" + Str(HandleGrupoUsuario),1,40)
  	SQL.ParamByName("PROTEGERREGISTRO").Value = CurrentQuery.FieldByName("PROTEGERREGISTRO").AsString

  	'Linhas abaixo incluídas SMS 9266
  	SQL.ParamByName("SFNFATURAAVULSA").Value = IIf(CurrentQuery.FieldByName("SFNFATURAAVULSA").IsNull, "N", CurrentQuery.FieldByName("SFNFATURAAVULSA").Value)
  	SQL.ParamByName("SFNFATURABAIXAR").Value = IIf(CurrentQuery.FieldByName("SFNFATURABAIXAR").IsNull, "N", CurrentQuery.FieldByName("SFNFATURABAIXAR").Value)
  	SQL.ParamByName("SFNFATURACANCELAR").Value = IIf(CurrentQuery.FieldByName("SFNFATURACANCELAR").IsNull, "N", CurrentQuery.FieldByName("SFNFATURACANCELAR").Value)
  	SQL.ParamByName("SFNDOCUMENTOGERAR").Value = IIf(CurrentQuery.FieldByName("SFNDOCUMENTOGERAR").IsNull, "N", CurrentQuery.FieldByName("SFNDOCUMENTOGERAR").Value)
  	SQL.ParamByName("SFNDOCUMENTOBAIXAR").Value = IIf(CurrentQuery.FieldByName("SFNDOCUMENTOBAIXAR").IsNull, "N", CurrentQuery.FieldByName("SFNDOCUMENTOBAIXAR").Value)
  	SQL.ParamByName("SFNDOCUMENTOCANCELAR").Value = IIf(CurrentQuery.FieldByName("SFNDOCUMENTOCANCELAR").IsNull, "N", CurrentQuery.FieldByName("SFNDOCUMENTOCANCELAR").Value)
  	SQL.ParamByName("SFNDOCUMENTOIMPRIMIR").Value = IIf(CurrentQuery.FieldByName("SFNDOCUMENTOIMPRIMIR").IsNull, "N", CurrentQuery.FieldByName("SFNDOCUMENTOIMPRIMIR").Value)
  	SQL.ParamByName("SFNPARCELAMENTO").Value = IIf(CurrentQuery.FieldByName("SFNPARCELAMENTO").IsNull, "N", CurrentQuery.FieldByName("SFNPARCELAMENTO").Value)

  	SQL.ParamByName("ALTERAR").Value = CurrentQuery.FieldByName("ALTERAR").Value
  	SQL.ParamByName("INCLUIR").Value = CurrentQuery.FieldByName("INCLUIR").Value
  	SQL.ParamByName("EXCLUIR").Value = CurrentQuery.FieldByName("EXCLUIR").Value
  	SQL.ParamByName("ALTERARSENHA").Value = CurrentQuery.FieldByName("ALTERARSENHA").Value
  	SQL.ParamByName("INATIVO").Value = CurrentQuery.FieldByName("INATIVO").Value

  	SQL.ParamByName("DESENVOLVER").Value = CurrentQuery.FieldByName("DESENVOLVER").Value
  	SQL.ParamByName("USUARIOREMOTO").Value = CurrentQuery.FieldByName("USUARIOREMOTO").Value
  	SQL.ParamByName("RETERSENHA").Value = CurrentQuery.FieldByName("RETERSENHA").Value
  	SQL.ParamByName("PERMITESOBREPORHORARIO").Value = CurrentQuery.FieldByName("PERMITESOBREPORHORARIO").Value
  	SQL.ParamByName("EMAIL").Value = CurrentQuery.FieldByName("EMAIL").AsString
  	SQL.ParamByName("ESPECIALIDADES").Value = CurrentQuery.FieldByName("ESPECIALIDADES").Value

  	'
  	If Not CurrentQuery.FieldByName("FILIALPADRAO").IsNull Then SQL.ParamByName("FILIALPADRAO").Value = CurrentQuery.FieldByName("FILIALPADRAO").AsInteger
  	If Not CurrentQuery.FieldByName("IDIOMA").IsNull Then SQL.ParamByName("IDIOMA").Value = CurrentQuery.FieldByName("IDIOMA").AsInteger
  	If Not CurrentQuery.FieldByName("TIPOAUTORIZACAOPADRAO").IsNull Then SQL.ParamByName("TIPOAUTORIZACAOPADRAO").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAOPADRAO").AsInteger

  	If Not CurrentQuery.FieldByName("ULTIMAEMPRESA").IsNull Then SQL.ParamByName("ULTIMAEMPRESA").Value = CurrentQuery.FieldByName("ULTIMAEMPRESA").AsInteger
  	If Not CurrentQuery.FieldByName("ULTIMAFILIAL").IsNull Then SQL.ParamByName("ULTIMAFILIAL").Value = CurrentQuery.FieldByName("ULTIMAFILIAL").AsInteger
  	If Not CurrentQuery.FieldByName("ULTIMOMODULO").IsNull Then SQL.ParamByName("ULTIMOMODULO").Value = CurrentQuery.FieldByName("ULTIMOMODULO").AsInteger
  	If Not CurrentQuery.FieldByName("FORMWIDTH").IsNull Then SQL.ParamByName("FORMWIDTH").Value = CurrentQuery.FieldByName("FORMWIDTH").AsInteger
  	If Not CurrentQuery.FieldByName("FORMHEIGHT").IsNull Then SQL.ParamByName("FORMHEIGHT").Value = CurrentQuery.FieldByName("FORMHEIGHT").AsInteger
  	If Not CurrentQuery.FieldByName("MUDANCAFASE").IsNull Then SQL.ParamByName("MUDANCAFASE").Value = CurrentQuery.FieldByName("MUDANCAFASE").AsInteger
  	If Not CurrentQuery.FieldByName("SMTPPORT").IsNull Then SQL.ParamByName("SMTPPORT").Value = CurrentQuery.FieldByName("SMTPPORT").AsString
  	If Not CurrentQuery.FieldByName("SMTPSERVER").IsNull Then SQL.ParamByName("SMTPSERVER").Value = CurrentQuery.FieldByName("SMTPSERVER").AsString
  	'If Not CurrentQuery.FieldByName("MEMO").IsNull Then SQL.ParamByName("MEMO").Value = CurrentQuery.FieldByName("MEMO").Value


  	SQL.ParamByName("MEMO").Value = CurrentQuery.FieldByName("MEMO").AsString

  	SQL.ParamByName("UNIDADEMEDIDA").Value = CurrentQuery.FieldByName("UNIDADEMEDIDA").AsInteger
  	SQL.ParamByName("VIEWGRID").Value = CurrentQuery.FieldByName("VIEWGRID").AsString

  	SQL.ExecSQL

  	'CRIA AGENDA
  	SQL1.Add("SELECT * FROM Z_USUARIOAGENDAS WHERE USUARIO = " + CurrentQuery.FieldByName("HaNdle").AsString)
  	SQL1.Active = True

  	SQL.Clear
  	SQL.Add("INSERT INTO Z_USUARIOAGENDAS (HANDLE, USUARIO, AGENDA) VALUES (:HANDLE, :USUARIO, :AGENDA)")

  	While Not SQL1.EOF
	    SQL.Active = False
	    SQL.ParamByName("HANDLE").Value = NewHandle("Z_USUARIOAGENDAS")
	    SQL.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL.ParamByName("AGENDA").Value = SQL1.FieldByName("AGENDA").AsInteger

	    SQL.ExecSQL
	    SQL1.Next

  	Wend

  	'CRIA EMPRESA
  	SQL1.Clear
  	SQL1.Add("SELECT * FROM Z_GRUPOUSUARIOEMPRESAS WHERE USUARIO = " + CurrentQuery.FieldByName("HaNdle").AsString)
  	SQL1.Active = True


  	While Not SQL1.EOF
	    SQL.Clear
	    SQL.Add("INSERT INTO Z_GRUPOUSUARIOEMPRESAS (HANDLE, GRUPO, USUARIO, EMPRESA) VALUES (:HANDLE, :GRUPO, :USUARIO, :EMPRESA)")

	    HandleGrupoUsuarioEmpresa = NewHandle("Z_GRUPOUSUARIOEMPRESAS")

	    SQL.ParamByName("HANDLE").Value = HandleGrupoUsuarioEmpresa
	    SQL.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL.ParamByName("GRUPO").Value = SQL1.FieldByName("GRUPO").AsInteger
	    SQL.ParamByName("EMPRESA").Value = SQL1.FieldByName("EMPRESA").AsInteger

	    SQL.ExecSQL

    	SQL2.Clear
    	SQL2.Add("SELECT *")
    	SQL2.Add("  FROM Z_GRUPOUSUARIOEMPRESAFILIAIS")
    	SQL2.Add("  WHERE EMPRESA = :E")
    	SQL2.Add("    AND USUARIO = :U")
    	SQL2.ParamByName("E").Value = SQL1.FieldByName("EMPRESA").AsInteger
    	SQL2.ParamByName("U").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL2.Active = True

    	While Not SQL2.EOF
      	SQL.Clear
      	SQL.Add("INSERT INTO Z_GRUPOUSUARIOEMPRESAFILIAIS (HANDLE, GRUPO, USUARIO, EMPRESA, FILIAL,SOMENTELEITURA) VALUES (:HANDLE, :GRUPO, :USUARIO, :EMPRESA, :FILIAL, :SOMENTELEITURA)")

      	SQL.ParamByName("HANDLE").Value = NewHandle("Z_GRUPOUSUARIOEMPRESAFILIAIS")
      	SQL.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("GRUPO").AsInteger
      	SQL.ParamByName("USUARIO").Value = HandleGrupoUsuario
      	SQL.ParamByName("EMPRESA").Value = SQL2.FieldByName("EMPRESA").AsInteger
      	SQL.ParamByName("FILIAL").Value = SQL2.FieldByName("FILIAL").AsInteger
      	SQL.ParamByName("SOMENTELEITURA").Value = SQL2.FieldByName("SOMENTELEITURA").AsString

      	SQL.ExecSQL
      	SQL2.Next
	    Wend

	    SQL1.Next

  	Wend

  	'CRIA NEGACOES
  	SQL1.Clear
  	SQL1.Add("SELECT * FROM SAM_USUARIO_MOTIVONEGACAO WHERE USUARIO = " + CurrentQuery.FieldByName("HaNdle").AsString)
  	SQL1.Active = True

  	SQL.Clear
  	SQL.Add("INSERT INTO SAM_USUARIO_MOTIVONEGACAO (HANDLE, USUARIO, MOTIVONEGACAO) VALUES (:HANDLE, :USUARIO, :MOTIVONEGACAO)")

  	While Not SQL1.EOF
	    SQL.Active = False
	    SQL.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVONEGACAO")
	    SQL.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL.ParamByName("MOTIVONEGACAO").Value = SQL1.FieldByName("MOTIVONEGACAO").AsInteger

	    SQL.ExecSQL
	    SQL1.Next

  	Wend

  	'CRIA NIVEIS DE AUTORIZA#ÇO
  	'SQL1.Clear
  	'SQL1.Add("SELECT * FROM SAM_GRUPOUSUARIO WHERE USUARIO = "+ CurrentQuery.FieldByName("HaNdle").AsString)
  	'SQL1.Active = True

  	'SQL.Clear
  	'SQL.Add("INSERT INTO SAM_GRUPOUSUARIO (HANDLE, USUARIO, NIVELAUTORIZACAO) VALUES (:HANDLE, :USUARIO, :NIVELAUTORIZACAO)")

  	'While Not SQL1.EOF
  	'SQL.Active = False
  	'	SQL.ParamByName("HANDLE").Value = NewHandle("SAM_GRUPOUSUARIO")
  	'	SQL.ParamByName("USUARIO").Value = HandleGrupoUsuario
  	'	SQL.ParamByName("NIVELAUTORIZACAO").Value = SQL1.FieldByName("NIVELAUTORIZACAO").AsInteger

  	'	SQL.ExecSQL
  	'	SQL1.Next

  	'Wend

  	'Adicionar Permissões na Cópia
  	SQL1.Clear
  	SQL1.Add("INSERT INTO Z_GRUPOUSUARIOGRUPOS (HANDLE, GRUPO, USUARIO, GRUPOADICIONADO) VALUES (:HANDLE, :GRUPO, :USUARIO, :GRUPOADICIONADO)")

  	SQL2.Clear
  	SQL2.Add("SELECT * FROM Z_GRUPOUSUARIOGRUPOS WHERE USUARIO= " + CurrentQuery.FieldByName("HANDLE").AsString)
  	SQL2.Active = True
  	While Not SQL2.EOF
	    SQL1.Active = False
	    SQL1.ParamByName("HANDLE").Value = NewHandle("Z_GRUPOUSUARIOGRUPOS")
	    SQL1.ParamByName("GRUPO").Value = SQL2.FieldByName("GRUPO").Value
	    SQL1.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL1.ParamByName("GRUPOADICIONADO").Value = SQL2.FieldByName("GRUPOADICIONADO").Value
	    SQL1.ExecSQL
	    SQL2.Next
  	Wend

  	'Filiais de Acesso na Cópia
  	SQL1.Clear
  	SQL1.Add("INSERT INTO Z_GRUPOUSUARIOS_FILIAIS (HANDLE, USUARIO, FILIAL, EXCLUIR, ALTERAR, INCLUIR)")
  	SQL1.Add("                             VALUES (:HANDLE, :USUARIO, :FILIAL, :EXCLUIR, :ALTERAR, :INCLUIR)")

  	SQL2.Clear
  	SQL2.Add("SELECT * FROM Z_GRUPOUSUARIOS_FILIAIS WHERE USUARIO= " + CurrentQuery.FieldByName("HANDLE").AsString)
  	SQL2.Active = True
  	While Not SQL2.EOF
	    SQL1.Active = False
	    SQL1.ParamByName("HANDLE").Value = NewHandle("Z_GRUPOUSUARIOS_FILIAIS")
	    SQL1.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL1.ParamByName("FILIAL").Value = SQL2.FieldByName("FILIAL").Value
	    SQL1.ParamByName("EXCLUIR").Value = SQL2.FieldByName("EXCLUIR").Value
	    SQL1.ParamByName("ALTERAR").Value = SQL2.FieldByName("ALTERAR").Value
	    SQL1.ParamByName("INCLUIR").Value = SQL2.FieldByName("INCLUIR").Value
	    SQL1.ExecSQL
	    SQL2.Next
  	Wend

  	'Níveis de Autorização na Cópia
  	SQL1.Clear
  	SQL1.Add("INSERT INTO SAM_GRUPOUSUARIO (HANDLE, USUARIO, NIVELAUTORIZACAO, NIVEL) VALUES (:HANDLE, :USUARIO, :NIVELAUTORIZACAO, :NIVEL)")

  	SQL2.Clear
  	SQL2.Add("SELECT * FROM SAM_GRUPOUSUARIO WHERE USUARIO= " + CurrentQuery.FieldByName("HANDLE").AsString)
  	SQL2.Active = True
  	While Not SQL2.EOF
	    SQL1.Active = False
	    SQL1.ParamByName("HANDLE").Value = NewHandle("SAM_GRUPOUSUARIO")
	    SQL1.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL1.ParamByName("NIVELAUTORIZACAO").Value = SQL2.FieldByName("NIVELAUTORIZACAO").Value
	    SQL1.ParamByName("NIVEL").Value = SQL2.FieldByName("NIVEL").Value
	    SQL1.ExecSQL
	    SQL2.Next
  	Wend


  	'MOTIVO GLOSA
  	SQL1.Clear
  	SQL1.Add("INSERT INTO SAM_USUARIO_MOTIVOGLOSA")
  	SQL1.Add("(HANDLE,MOTIVOGLOSA,USUARIO)")
  	SQL1.Add("VALUES")
	  	SQL1.Add("(:HANDLE,:MOTIVOGLOSA,:USUARIO)")

  	SQL2.Clear
  	SQL2.Add("SELECT * FROM SAM_USUARIO_MOTIVOGLOSA WHERE USUARIO =" + CurrentQuery.FieldByName("HANDLE").AsString)
  	SQL2.Active = True

  	While Not SQL2.EOF
	    SQL1.Active = False
	    SQL1.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVOGLOSA")
	    SQL1.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL1.ParamByName("MOTIVOGLOSA").Value = SQL2.FieldByName("MOTIVOGLOSA").AsInteger
	    SQL1.ExecSQL
	    SQL2.Next
  	Wend

  	'Agenda

  	SQL1.Clear
  	SQL1.Add("INSERT INTO Z_AGENDAUSUARIOS")
  	SQL1.Add("(HANDLE,USUARIO,AGENDA,ALERTAR,ANTECEDENCIA)")
  	SQL1.Add("VALUES")
  	SQL1.Add("(:HANDLE,:USUARIO,:AGENDA,:ALERTAR,:ANTECEDENCIA)")

  	SQL2.Clear
  	SQL2.Add("SELECT * FROM Z_AGENDAUSUARIOS WHERE USUARIO =" + CurrentQuery.FieldByName("HANDLE").AsString)
  	SQL2.Active = True

  	While Not SQL2.EOF
	    SQL1.Active = False
	    SQL1.ParamByName("HANDLE").Value = NewHandle("Z_AGENDAUSUARIOS")
	    SQL1.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL1.ParamByName("AGENDA").Value = SQL2.FieldByName("AGENDA").AsInteger
	    SQL1.ParamByName("ANTECEDENCIA").Value = SQL2.FieldByName("ANTECEDENCIA").AsInteger
	    SQL1.ParamByName("ALERTAR").Value = SQL2.FieldByName("ALERTAR").AsInteger
	    SQL1.ExecSQL
	    SQL2.Next
  	Wend

  	'ALCADAS
  	SQL1.Clear
  	SQL1.Add("INSERT INTO Z_GRUPOUSUARIOALCADAS")
  	SQL1.Add("(HANDLE,USUARIO,ALCADA,LIMITE)")
  	SQL1.Add("VALUES")
  	SQL1.Add("(:HANDLE,:USUARIO,:ALCADA,:LIMITE)")

  	SQL2.Clear
  	SQL2.Add("SELECT * FROM Z_GRUPOUSUARIOALCADAS WHERE USUARIO =" + CurrentQuery.FieldByName("HANDLE").AsString)
  	SQL2.Active = True

  	While Not SQL2.EOF
	    SQL1.Active = False
	    SQL1.ParamByName("HANDLE").Value = NewHandle("Z_GRUPOUSUARIOALCADAS")
	    SQL1.ParamByName("USUARIO").Value = HandleGrupoUsuario
	    SQL1.ParamByName("ALCADA").Value = SQL2.FieldByName("ALCADA").AsInteger
	    SQL1.ParamByName("LIMITE").Value = SQL2.FieldByName("LIMITE").AsFloat
	    SQL1.ExecSQL
	    SQL2.Next
	    lPassou = True
  	Wend

  	Set SQL = Nothing
  	Set SQL1 = Nothing
  	Set SQL2 = Nothing

  	Commit
  	If lPassou Then
	    bsShowMessage("Todos os registros foram copiados com êxito ! " + Chr(13) + "Mas você deverá escolher um recurso da clínica para este usuário !", "I")
  	End If

  	RefreshNodesWithTable "Z_GRUPOUSUARIOS"
  	Exit Sub
	FIM:
  	Rollback
  	bsShowMessage("Não foi possível copiar o Usuário :" + Str(Error), "E")
  	Set SQL = Nothing
  	Set SQL1 = Nothing
	  Set SQL2 = Nothing

  End If

End Sub

Public Sub COPIAR_OnClick()
  Dim ObjCopy As Object
  Set ObjCopy = CreateBennerObject("CS.Security")
  ObjCopy.CopySec(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentSystem)
  Set ObjCopy = Nothing
End Sub

Public Sub TABLE_AfterEdit()
  iGrupo = CurrentQuery.FieldByName("GRUPO").AsInteger
End Sub


Public Sub TABLE_AfterInsert()
  iGrupo = 0
End Sub


Public Sub TABLE_AfterPost()
  If (iGrupo > 0) And (CurrentQuery.FieldByName("GRUPO").AsInteger <> iGrupo) Then
    Dim qWork As Object
    Set qWork = NewQuery

    qWork.Add("DELETE FROM Z_GRUPOUSUARIOGRUPOS WHERE GRUPO = " + CStr(iGrupo) + _
              " AND USUARIO = " + CurrentQuery.FieldByName("HANDLE").AsString)
    qWork.ExecSQL
    qWork.Clear
    qWork.Add("UPDATE Z_GRUPOUSUARIOEMPRESAS SET GRUPO = " + CurrentQuery.FieldByName("GRUPO").AsString + _
              " WHERE USUARIO = " + CurrentQuery.FieldByName("HANDLE").AsString)
    qWork.ExecSQL
    qWork.Clear
    qWork.Add("UPDATE Z_GRUPOUSUARIOEMPRESAFILIAIS SET GRUPO = " + CurrentQuery.FieldByName("GRUPO").AsString + _
              " WHERE USUARIO = " + CurrentQuery.FieldByName("HANDLE").AsString)
    qWork.ExecSQL

    Set qWork = Nothing
  End If

  'Em modo Web não está disponíve a edição do campo "Senha". Desta forma, na inclusão de um novo usuário
  'será gerada uma senha aleatória e a mesma será enviada para o email do usuário
  If WebMode And _
     vsModoEdicao = "I" Then
    On Error GoTo erro

      Dim Usuario  As CSSystemUser
      Set Usuario = NewSystemUser

      Usuario.SelectUser(CurrentQuery.FieldByName("HANDLE").AsInteger)
      Usuario.ChangePasswordAndMail("Sistema Benner Saúde: Envio da senha de acesso ao sistema", "")
      Set Usuario = Nothing

      Exit Sub
    erro:
      bsShowMessage("Erro ao enviar senha para e-mail do usuário: " + Err.Description, "I")
  End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
	    TIPOAUTORIZACAOPADRAO.WebLocalWhere = "(DATAINICIAL <= " + CurrentSystem.SQLDate(CurrentSystem.ServerDate) + ") AND (DATAFINAL Is Null)" 'SMS 81867 - Débora Rebello - 21/05/2007
	ElseIf VisibleMode Then
		TIPOAUTORIZACAOPADRAO.LocalWhere = "(DATAINICIAL <= " + CurrentSystem.SQLDate(CurrentSystem.ServerDate) + ") AND (DATAFINAL Is Null)" 'SMS 81867 - Débora Rebello - 21/05/2007
		BOTAOLIBERARGLOSA.Visible = False
        BOTAOLIBERARNEGACAO.Visible = False
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vsModoEdicao = "A"
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  vsModoEdicao = "I"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As BPesquisa
  Dim SQL2 As BPesquisa
  Dim SQL3 As BPesquisa
  Set SQL = NewQuery
  Set SQL2 = NewQuery
  Set SQL3 = NewQuery

  SQL.Add("SELECT COUNT(1) QTD                                               ")
  SQL.Add("  FROM SAM_PRESTADOR P                                            ")
  SQL.Add("  JOIN SAM_CATEGORIA_PRESTADOR CP ON (P.CATEGORIA = CP.HANDLE)    ")
  SQL.Add(" WHERE P.CPFCNPJ LIKE :CPFCNPJ AND CP.BLOQUEARINCLUSAOBENEF = 'S' ")

  SQL.ParamByName("CPFCNPJ").AsString = CurrentQuery.FieldByName("CPF").AsString
  SQL.Active = True

  SQL2.Add("SELECT CRITICAINCPRESTADOR ")
  SQL2.Add("  FROM Z_GRUPOS G          ")
  SQL2.Add(" WHERE G.HANDLE = :GRUPO   ")

  SQL2.ParamByName("GRUPO").AsString = CurrentQuery.FieldByName("GRUPO").AsString
  SQL2.Active = True

  If (Not(CurrentQuery.FieldByName("CPF").IsNull)) Then

	  If Not(IsValidCPF(CurrentQuery.FieldByName("CPF").AsString)) Then
	    bsShowMessage("O CPF informado não é válido", "E")
		CanContinue = False
	  ElseIf ((SQL2.FieldByName("CRITICAINCPRESTADOR").AsString = "S") And (SQL.FieldByName("QTD").AsInteger > 0)) And (CurrentQuery.FieldByName("INATIVO").AsString = "N") Then
		If bsShowMessage("O CPF informado é semelhante ao de um prestador do sistema. Deseja continuar mesmo assim?", "Q") = vbNo Then
		  If VisibleMode Then
		    CanContinue = False
		  End If
		  Exit Sub
		End If
	  End If

  End If

  If WebMode And _
     (CurrentQuery.FieldByName("EMAIL").IsNull Or _
      Trim(CurrentQuery.FieldByName("EMAIL").AsString) = "") Then
    CanContinue = False
    bsShowMessage("E-mail deve ser preenchido!", "E")
  End If

  SQL3.Add("SELECT COUNT(1) QTD                     ")
  SQL3.Add("  FROM Z_GRUPOUSUARIOS                  ")
  SQL3.Add(" WHERE UPPER(APELIDO) = UPPER(:APELIDO) ")
  SQL3.Add("   AND HANDLE <> :HANDLE                ")

  SQL3.ParamByName("APELIDO").AsString = CurrentQuery.FieldByName("APELIDO").AsString
  SQL3.ParamByName("HANDLE").AsString = CurrentQuery.FieldByName("HANDLE").AsString
  SQL3.Active = True

  If SQL3.FieldByName("QTD").AsInteger > 0 Then
    CanContinue = False
    bsShowMessage("Já existe outro usuário com o mesmo apelido!", "E")
  End If

  Set SQL = Nothing
  Set SQL2 = Nothing
  Set SQL3 = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
      Case "COPIAR"
        COPIAR_OnClick
      Case "BSBOTAOCOPIAR"
        BSBOTAOCOPIAR_OnClick
      Case "BOTAONOVASENHA"
        BOTAONOVASENHA_OnClick
      Case "BOTAOLIBERARGLOSA"
        BOTAOLIBERARGLOSA_OnClick
      Case "BOTAOLIBERARNEGACAO"
        BOTAOLIBERARNEGACAO_OnClick
  End Select
End Sub

Public Sub TABLE_UpdateRequired()
  If WebMode And _
     vsModoEdicao = "I" Then
    CurrentQuery.FieldByName("SENHA").AsString = "Benner"
  End If
End Sub
