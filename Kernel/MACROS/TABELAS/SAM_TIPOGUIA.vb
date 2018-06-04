'HASH: BED63D0272EEB2E1F7C26E5BA26BEC66
'MACRO:SAM_TIPOGUIA
'#Uses "*bsShowMessage"

Public Sub BOTAOEXCLUIRLEIAUTE_OnClick()
'SMS 101044 - Paulo Melo - 14/08/2008 - Esse nem era o propósito da sms, mas esse botão estava sem funcionalidade...

  If CurrentQuery.State = 3 Then
		bsShowMessage("Modelo sendo inserido. Nada a excluir.", "I")
  Else
	  Dim SQL2 As Object
	  Set SQL2 = NewQuery


	  Dim qGuia As Object
	  Set qGuia = NewQuery

	  SQL2.Clear
	  SQL2.Add("SELECT COUNT(1) QTD FROM SAM_PEG WHERE TIPODEGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.Active = True

	  If SQL2.FieldByName("QTD").AsInteger > 0 Then
		bsShowMessage("Este tipo de guia não pode ser excluído, pois esta associado a um PEG!", "E")
		Set SQL2  = Nothing
  	  	Set qGuia = Nothing
	  	Exit Sub
	  End If

	  qGuia.Clear
	  qGuia.Add("SELECT COUNT(1) QTD FROM SAM_GUIA WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  qGuia.Active = True

	  If qGuia.FieldByName("QTD").AsInteger > 0 Then
		bsShowMessage("Este modelo não pode ser excluído, pois esta associado a uma guia!", "E")
		Set SQL2  = Nothing
  	  	Set qGuia = Nothing
	  	Exit Sub
	  End If

	  qGuia.Clear
	  qGuia.Add("SELECT MODELOGUIASUS FROM SAM_PARAMETROSPROCCONTAS")
	  qGuia.Active = True

	  SQL2.Clear
	  SQL2.Add("SELECT MG.HANDLE            							 ")
	  SQL2.Add("FROM SAM_TIPOGUIA_MDGUIA MG 						     ")
	  SQL2.Add("JOIN SAM_TIPOGUIA  TP ON (MG.TIPOGUIA = TP.HANDLE)       ")
	  SQL2.Add("WHERE TP.HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.Active = True

	  If qGuia.FieldByName("MODELOGUIASUS").AsInteger = SQL2.FieldByName("HANDLE").AsInteger Then
	  	bsShowMessage("Esse modelo de guia não pode ser excluído, pois está associado ao Modelo de Guia SUS, nos Parâmetros Gerais do Processamento de Contas", "E")
	  	Set SQL2  = Nothing
  	    Set qGuia = Nothing
	  	Exit Sub
	  End If

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_GUIA WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_EVENTO WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_MATMED WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_GRAU WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_EVENTOTGE WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_TIPOTRAT WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_REGATEND WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_OBJTRAT WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_LOCALATEND WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_FINATEND WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_CONDATEND WHERE MODELOGUIA IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA WHERE TIPOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA WHERE HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  RefreshNodesWithTable("SAM_TIPOGUIA")

	  Set SQL2  = Nothing
  	  Set qGuia = Nothing
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not (VerificaExisteCodigoDespesas) Then
	bsShowMessage("Não é possível alterar o Tipo de Guia TISS. O campo 'Código de Outras Despesas' está inserido em algum " + _
		"Modelo de Guia desse Tipo de Guia e esse campo é exclusivo do Tipo de Guia TISS 'Outras Despesas'", "E")
	CanContinue = False
	Exit Sub
  End If
End Sub
Public Function VerificaExisteCodigoDespesas() As Boolean
  'SMS 81587 - Débora Rebello - 16/05/2007
  Dim SQL As Object
  Dim qCampo As Object
  Set SQL = NewQuery
  Set qCampo = NewQuery

  SQL.Active = False

  SQL.Clear

  SQL.Add("SELECT GC.NOMECAMPO2 ")
  SQL.Add("  FROM SAM_TIPOGUIA T ")
  SQL.Add("  JOIN  SAM_TIPOGUIA_MDGUIA        TM On (T.HANDLE = TM.TIPOGUIA)    ")
  SQL.Add("  JOIN SAM_TIPOGUIA_MDGUIA_EVENTO  TE On (TE.MODELOGUIA = TM.Handle) ")
  SQL.Add("  JOIN SIS_MODELOGUIA_CAMPOS       GC On (GC.HANDLE = TE.SISCAMPO)   ")
  SQL.Add(" WHERE T.HANDLE = :TIPOGUIA ")
  SQL.Add("       AND GC.NOMECAMPO2 = :CAMPOEVENTO ")

  SQL.ParamByName("TIPOGUIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("CAMPOEVENTO").AsString = "EventoCodigoDespRealizadas"
  SQL.Active = True

  If (Not SQL.EOF) Then
	'se o campo EventoCodigoDespRealizadas estiver em algum modelo de guia desse tipo de guia,
	'só poderá ser salvo se for do tipo de guia TISS "Outras Despesas"
	If (CurrentQuery.FieldByName("TIPOGUIATISS").AsString = "D") Then
	  VerificaExisteCodigoDespesas = True
	Else
	  VerificaExisteCodigoDespesas = False
	End If
  Else
	VerificaExisteCodigoDespesas = True
  End If

  Set SQL = Nothing
  Set qCampo = Nothing
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOEXCLUIRLEIAUTE" Then
		BOTAOEXCLUIRLEIAUTE_OnClick
	End If
End Sub
