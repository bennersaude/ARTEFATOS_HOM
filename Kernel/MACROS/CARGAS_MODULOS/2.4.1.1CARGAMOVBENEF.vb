'HASH: 541E38057F01A7414AFB8CED9E952047
 

Public Sub BOTAOMOVIMENTACAO_OnClick()

  Dim Interface As Object

  Set Interface = CreateBennerObject("BENNER.SAUDE.DESKTOP.BENEFICIARIOS.MONITORANALISEMOVIMENTACOES.Rotinas")
  Interface.ExecutaMonitor(CurrentSystem)
  Set Interface = Nothing

End Sub

Public Sub BOTAORECUSARREGISTRO_OnClick()

	Dim vsSqlParamGerais As Object
	Dim vsSqlUpdate As Object

	Set vsSqlParamGerais = NewQuery
	Set vsSqlUpdate = NewQuery

	vsSqlParamGerais.Clear
	vsSqlParamGerais.Active = False
	vsSqlParamGerais.Add("SELECT PRAZOENVIODOCUMENTOS FROM SAM_PARAMETROSWEB")
    vsSqlParamGerais.Active = True

	vsSqlUpdate.Clear
	vsSqlUpdate.Add("UPDATE WEB_SAM_BENEFICIARIO SET SITUACAO = '3' WHERE (SITUACAO = '1' OR SITUACAO = '5') AND DATAHORAINCLUSAO + :DIAS < :DATAFINAL")
	vsSqlUpdate.ParamByName("DIAS").AsInteger = vsSqlParamGerais.FieldByName("PRAZOENVIODOCUMENTOS").AsInteger
	vsSqlUpdate.ParamByName("DATAFINAL").AsDateTime = ServerNow
	vsSqlUpdate.ExecSQL

	MsgBox("Registros rejeitados com sucesso!")

	RefreshNodesWithTable("WEB_SAM_BENEFICIARIO")

	Set vsSqlParamGerais = Nothing
	Set vsSqlUpdate      = Nothing

End Sub
