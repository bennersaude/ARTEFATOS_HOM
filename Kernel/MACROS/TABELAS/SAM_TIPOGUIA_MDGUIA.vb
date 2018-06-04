'HASH: D03B3D94E65A03D9EE1D78BACADD0E8B
'#Uses "*bsShowMessage"

Public Sub DeleteModelo(CanContinue As Boolean)

  If CurrentQuery.State = 3 Then
	  bsShowMessage("Modelo sendo inserido. Nada a excluir.", "E")
	  CanContinue = False
  Else
	  Dim SQL2 As Object
	  Set SQL2 = NewQuery

	  SQL2.Clear
	  SQL2.Add("SELECT COUNT(1) QTD FROM SAM_GUIA WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.Active = True

	  If SQL2.FieldByName("QTD").AsInteger > 0 Then
		bsShowMessage("Este modelo não pode ser excluído, pois esta associado a uma guia!", "E")
		CanContinue = False
		Set SQL2  = Nothing
	  	Exit Sub
	  End If

	  SQL2.Clear
	  SQL2.Add("SELECT MODELOGUIASUS FROM SAM_PARAMETROSPROCCONTAS")
	  SQL2.Active = True

	  If SQL2.FieldByName("MODELOGUIASUS").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger Then
	  	bsShowMessage("Esse modelo de guia não pode ser excluído, pois está associado ao Modelo de Guia SUS, nos Parâmetros Gerais do Processamento de Contas", "E")
	  	CanContinue = False
	  	Set SQL2  = Nothing
	  	Exit Sub
	  End If

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_GUIA WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_EVENTO WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_MATMED WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_GRAU WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_EVENTOTGE WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_TIPOTRAT WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_REGATEND WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_OBJTRAT WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_LOCALATEND WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

      SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_FINATEND WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  SQL2.Clear
	  SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA_CONDATEND WHERE MODELOGUIA = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  SQL2.ExecSQL

	  If VisibleMode Then
	  	SQL2.Clear
	  	SQL2.Add("DELETE FROM SAM_TIPOGUIA_MDGUIA WHERE HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString)
	  	SQL2.ExecSQL
	  End If

	  RefreshNodesWithTable "SAM_TIPOGUIA_MDGUIA"

	  Set SQL2  = Nothing
  End If
End Sub

Public Sub BOTAODUPLICARLEIAUTE_OnClick()
  Dim INTERFACE As Object
  Dim Legenda As String
  Dim HandleCarga As Long
  Dim Tabela As String
  Dim CurrentHandle As Long
  Set INTERFACE = CreateBennerObject("SamPegDigit.Digitacao")

  INTERFACE.DuplicarModeloGuia(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("DESCRICAO").AsString)

  Set INTERFACE = Nothing

  RefreshNodesWithTable "SAM_TIPOGUIA_MDGUIA"
End Sub

Public Sub BOTAOVALIDARTISS_OnClick()
  Dim vModeloTISSOk As String
  Dim vCampoEhObrigatorio As String
  Dim vsMsg As String
  Dim INTERFACE As Object
  Set INTERFACE = CreateBennerObject("BsPro006.Geral")

  vsMsg = INTERFACE.VerificaModeloTISS(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "", vModeloTISSOk, vCampoEhObrigatorio)

  Set INTERFACE = Nothing

  If vModeloTISSOk= "S" Then
	bsShowMessage("Foram definidos todos os campos obrigatórios do Modelo TISS corretamente.", "I")
  Else
    If vsMsg <> "" Then
      bsShowMessage(vsMsg, "I")
    End If
  End If
End Sub

Public Sub TABLE_AfterScroll()
  BOTAOGERAREVENTOS.Visible=False
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear

  SQL.Add("SELECT TABTIPOGUIA, TIPOGUIATISS")
  SQL.Add("  FROM SAM_TIPOGUIA")
  SQL.Add(" WHERE HANDLE = :HANDLE")

  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOGUIA").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TABTIPOGUIA").AsInteger = 3 Then
	EXIGEHISTORICOBUCAL.Visible = True
  Else
	EXIGEHISTORICOBUCAL.Visible = False
  End If

  'sms 77454
  If (SQL.FieldByName("TIPOGUIATISS").AsString = "N") Or (SQL.FieldByName("TIPOGUIATISS").IsNull) Then
	BOTAOVALIDARTISS.Visible = False
  Else
	BOTAOVALIDARTISS.Visible = True
  End If
End Sub

Public Sub BOTAOEXCLUIRLEIAUTE_OnClick()
	Dim CanContinue As Boolean
	DeleteModelo(CanContinue)
End Sub

Public Sub BOTAOPREVERFORM_OnClick()
  Dim INTERFACE As Object

  Set INTERFACE = CreateBennerObject("BSPRO006.ROTINAS")

  INTERFACE.PreverFormulario(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set INTERFACE = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	DeleteModelo(CanContinue)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL3 As Object
  Dim q1 As Object
  Dim q2 As Object
  Dim q4 As Object
  Dim primeiromodelo As Long
  Dim processaPrimeiroModelo As Boolean
  Set SQL3 = NewQuery
  Set q1 = NewQuery
  Set q2 = NewQuery
  Set q4 = NewQuery

  CurrentQuery.UpdateRecord

  If Not InTransaction Then StartTransaction

  SQL3.Clear

  SQL3.Add("SELECT Count(*) VALOR FROM SAM_TIPOGUIA_MDGUIA ")
  SQL3.Add(" WHERE TIPOGUIA = " + CurrentQuery.FieldByName("TIPOGUIA").AsString)
  SQL3.Add("   AND PADRAO = 'S'")
  SQL3.Add("   AND HANDLE <> " + CurrentQuery.FieldByName("HANDLE").AsString)

  SQL3.Active = True

  q4.Clear

  q4.Add("SELECT COUNT(1) QTDE         ")
  q4.Add("  FROM SAM_TIPOGUIA_MDGUIA   ")
  q4.Add(" WHERE (CODIGO = :CODIGO)    ")
  q4.Add("   AND HANDLE <> :HANDLE     ")
  q4.Add("   AND (HANDLE IN (SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA))")

  q4.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  q4.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  q4.Active = True

  If (q4.FieldByName("QTDE").AsInteger > 0) Then
	bsShowMessage("Já existe um modelo de guia com este código", "I")
	CanContinue = False
	Exit Sub
  End If

  ' seleciona o primeiro modelo de guia que existe para o tipo de guia corrente '
  ' sendo utilizado na marcação do tipo de guia padrão caso nenhum modelo de guia esteja selecinado como padrão '
  q2.Clear

  q2.Add("SELECT HANDLE, PADRAO FROM SAM_TIPOGUIA_MDGUIA")
  q2.Add("WHERE TIPOGUIA = " + CurrentQuery.FieldByName("TIPOGUIA").AsString)
  q2.Add("ORDER BY DESCRICAO")

  q2.Active = True

  primeiromodelo = q2.FieldByName("HANDLE").AsInteger

  'Caso nenhum modelo de guia seja selecionado como padrao o primeiro modelo de guia selecionado atraves do '
  'SELECT acima é marcado como padrão'
  processaPrimeiroModelo = True

  If (CurrentQuery.FieldByName("PADRAO").AsString = "S") And (CurrentQuery.FieldByName ("REGIMEPAGTO").AsString = "A") Then
	bsShowMessage("Essa alteração modificará o modelo de guia corrente fazendo com que seja " + _
		   "padrão para ambos (Credenciamento/Reembolso).", "I")
	q1.Clear

	q1.Add ("UPDATE sam_tipoguia_mdguia  ")
	q1.Add ("SET    padrao   =  'N'      ")
	q1.Add ("WHERE  tipoguia =  :tipoguia")
	q1.Add ("AND    padrao   =  'S'      ")
	q1.Add ("AND    handle   <> :handle  ")

	q1.ParamByName("TIPOGUIA").AsInteger = CurrentQuery.FieldByName("TIPOGUIA").AsInteger
	q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	q1.ExecSQL

	processaPrimeiroModelo = False
  End If

  If (CurrentQuery.FieldByName("PADRAO").AsString = "S") And (CurrentQuery.FieldByName ("REGIMEPAGTO").AsString = "R") Then
	q1.Clear

	q1.Add ("UPDATE sam_tipoguia_mdguia      ")
	q1.Add ("SET    padrao      =  'N'       ")
	q1.Add ("WHERE  tipoguia    =  :tipoguia ")
	q1.Add ("AND    padrao      =  'S'       ")
	q1.Add ("AND    regimepagto IN ('R', 'A')")
	q1.Add ("AND    handle      <> :handle   ")

	q1.ParamByName("TIPOGUIA").AsInteger = CurrentQuery.FieldByName("TIPOGUIA").AsInteger
	q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	q1.ExecSQL

	processaPrimeiroModelo = False
  End If

  If (CurrentQuery.FieldByName("PADRAO").AsString = "S") And (CurrentQuery.FieldByName ("REGIMEPAGTO").AsString = "C") Then
	q1.Clear

	q1.Add ("UPDATE sam_tipoguia_mdguia      ")
	q1.Add ("SET    padrao      =  'N'       ")
	q1.Add ("WHERE  tipoguia    =  :tipoguia ")
	q1.Add ("AND    padrao      =  'S'       ")
	q1.Add ("AND    regimepagto IN ('C', 'A')")
	q1.Add ("AND    handle      <> :handle   ")

	q1.ParamByName("TIPOGUIA").AsInteger = CurrentQuery.FieldByName("TIPOGUIA").AsInteger
	q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	q1.ExecSQL

	processaPrimeiroModelo = False
  End If

  If (processaPrimeiroModelo = True) And (SQL3.FieldByName("VALOR").AsInteger = 0) Then
	q1.Clear

	q1.Add("UPDATE SAM_TIPOGUIA_MDGUIA     ")
	q1.Add("SET    PADRAO = 'S'            ")
	q1.Add("WHERE  HANDLE = :PRIMEIROMODELO")

	q1.ParamByName("PRIMEIROMODELO").AsInteger = primeiromodelo

	q1.ExecSQL
  End If

  If CurrentQuery.FieldByName("HANDLE").AsInteger = primeiromodelo Then
	If q2.FieldByName("PADRAO").AsBoolean = True Then
	  If CurrentQuery.FieldByName("PADRAO").AsBoolean = False Then
		bsShowMessage("Não é permitido desmarcar o Padrão no primeiro modelo", "I")

		CurrentQuery.FieldByName("PADRAO").AsBoolean = True
	  End If
	End If
  End If

  If InTransaction Then Commit

  Set SQL3 = Nothing
  Set q1 = Nothing
  Set q2 = Nothing
  Set q4 = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
    Case "BOTAODUPLICARLEIAUTE"
      BOTAODUPLICARLEIAUTE_OnClick
    Case "BOTAOVALIDARTISS"
	  BOTAOVALIDARTISS_OnClick
	Case "BOTAOEXCLUIRLEIAUTE"
	  BOTAOEXCLUIRLEIAUTE_OnClick
	Case "BOTAOPREVERFORM"
	  BOTAOPREVERFORM_OnClick
  End Select
End Sub
