'HASH: 0FD0D8299C9F08A9ADDD4FA5BF4CD951
'macro SAM_LIVRO
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()

	Dim Msg As String
	Dim Result As String

	'If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		'bsShowMessage(Msg, "I")
		'CanContinue = False
		'Exit Sub
	'End If

	If bsShowMessage("Esta operação irá excluir todos os dados deste livro, " + (Chr(13)) + _
			   "inclusive seus encartes e rotinas. Deseja continua?", "Q") = vbYes Then

	    Dim Obj As Object

		'SMS 90283 - Ricardo Rocha - Adequacao WEB
		If VisibleMode Then
	    	Set Obj = CreateBennerObject("BSInterface0007.Rotinas")
	    	Obj.CancelarLivro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  		Else
	    	Dim vsMensagemErro As String
	    	Dim viRetorno As Long

	    	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                                     "BSPRE001", _
	                                     "Rotinas_CancelarLivro", _
	                                     "Cancelamento de Livro de Credenciados - Livro: " + _
	                                     CStr(CurrentQuery.FieldByName("HANDLE").AsInteger) + _
	                                     " Descrição: " + CurrentQuery.FieldByName("DESCRICAO").AsString, _
	                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
	                                     "SAM_LIVRO", _
	                                     "SITUACAOLIVRO", _
	                                     "", _
	                                     "", _
	                                     "C", _
	                                     False, _
	                                     vsMensagemErro, _
	                                     Null)

	    	If viRetorno = 0 Then
	      		bsShowMessage("Processo enviado para execução no servidor!", "I")
	    	Else
	      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	    	End If

  		End If
  		Set Obj = Nothing

  		RefreshNodesWithTable("SAM_LIVRO")
  	End If

End Sub

Public Sub BOTAOGERAR_OnClick()
	Dim Msg As String
	Dim Obj As Object

	'If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		'bsShowMessage(Msg, "I")
		'CanContinue = False
		'Exit Sub
	'End If

	'SMS 90283 - Ricardo Rocha - Adequacao WEB
	If VisibleMode Then
    	Set Obj = CreateBennerObject("BSInterface0007.Rotinas")
    	Obj.GerarLivro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  	Else
    	Dim vsMensagemErro As String
    	Dim viRetorno As Long

    	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSPRE001", _
                                     "Rotinas_GerarLivro", _
                                     "Geração de Livro de Credenciados - Livro: " + _
                                     CStr(CurrentQuery.FieldByName("HANDLE").AsInteger) + _
                                     " Descrição: " + CurrentQuery.FieldByName("DESCRICAO").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_LIVRO", _
                                     "SITUACAOLIVRO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     Null)

    	If viRetorno = 0 Then
      		bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If
  	End If
  	Set Obj = Nothing

	Dim SQL As Object
	BOTAOGERAR.Enabled = True
	BOTAOCANCELAR.Enabled = True
	BOTAOGERARENCARTE.Enabled = True

	If CurrentQuery.State = 3 Then
		CRIACAOUSUARIO.ReadOnly = False
	Else
		If Not CurrentQuery.FieldByName("CRIACAOUSUARIO").IsNull Then
			CRIACAOUSUARIO.ReadOnly = True
		End If
	End If

	If CurrentQuery.FieldByName("HANDLE").IsNull Then
		BOTAOGERAR.Enabled = False
		BOTAOGERARENCARTE.Enabled = False
		BOTAOCANCELAR.Enabled = False
	Else
		Set SQL = NewQuery

		SQL.Add("SELECT HANDLE FROM SAM_LIVRODADOS WHERE LIVRO = :LIVRO")

		SQL.ParamByName("LIVRO").Value = CurrentQuery.FieldByName("HANDLE").Value
		SQL.Active = True

		If SQL.FieldByName("HANDLE").IsNull Then
			BOTAOGERARENCARTE.Enabled = False
			BOTAOCANCELAR.Enabled = False
		Else
			BOTAOGERAR.Enabled = False
		End If
		Set SQL = Nothing

	End If
	RefreshNodesWithTable("SAM_LIVRO")

End Sub

Public Sub BOTAOGERARENCARTE_OnClick()
	Dim Obj As Object
	Dim Msg As String

	'If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		'bsShowMessage(Msg, "I")
		'CanContinue = False
		'Exit Sub
	'End If

	'SMS 90283 - Ricardo Rocha - Adequacao WEB
	If VisibleMode Then
		Set Obj = CreateBennerObject("BSInterface0007.Rotinas")
		Obj.GerarEncarte(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
	Else
		Dim vsMensagemErro As String
		Dim viRetorno As Long
		Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSPRE001", _
                                     "Rotinas_GerarEncarte", _
                                     "Geração de Encarte de Livro de Credenciados - Livro: " + _
                                     CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_LIVRO", _
                                     "SITUACAOENCARTE", _
                                     "", _
                                     "", _
                                     "P", _
                                     True, _
                                     vsMensagemErro, _
                                     Null)

    	If viRetorno = 0 Then
      		bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If
    End If
	Set Obj = Nothing

	RefreshNodesWithTable("SAM_LIVRO")
End Sub

Public Sub TABLE_AfterScroll()

	BOTAOGERAR.Enabled = True
	BOTAOGERARENCARTE.Enabled = True
	BOTAOCANCELAR.Enabled = True

	If CurrentQuery.State = 3 Then
		CRIACAOUSUARIO.ReadOnly = False
		CurrentQuery.FieldByName("CRIACAODATA").AsDateTime = ServerNow
	Else
		If Not CurrentQuery.FieldByName("CRIACAOUSUARIO").IsNull Then
			CRIACAOUSUARIO.ReadOnly = True
		End If
	End If

	If CurrentQuery.FieldByName("HANDLE").IsNull Then
		BOTAOGERAR.Enabled = False
		BOTAOCANCELAR.Enabled = False
		BOTAOGERARENCARTE.Enabled = False
	Else
        Dim SQL As Object
		Set SQL = NewQuery

		SQL.Add("SELECT HANDLE FROM SAM_LIVRODADOS WHERE LIVRO = :LIVRO")

		SQL.ParamByName("LIVRO").Value = CurrentQuery.FieldByName("HANDLE").Value
		SQL.Active = True

		If SQL.FieldByName("HANDLE").IsNull Then
			BOTAOCANCELAR.Enabled = False
			BOTAOGERARENCARTE.Enabled = False
		Else
			BOTAOGERAR.Enabled = False
		End If
		Set SQL = Nothing
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If NodeInternalCode <> 802 Then
    	VerificaSeProcessada(CanContinue)
  	End If

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

    If NodeInternalCode <> 802 Then
    	VerificaSeProcessada(CanContinue)
    End If

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	BOTAOGERAR.Enabled = True
End Sub

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("CRIACAOUSUARIO").Value = CurrentUser
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
		Case "BOTAOGERARENCARTE"
			BOTAOGERARENCARTE_OnClick
	End Select
End Sub

Public Sub VerificaSeProcessada(CanContinue As Boolean)
  Dim SQLRotFin As Object
  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT SITUACAOLIVRO FROM SAM_LIVRO WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQLRotFin.Active = True
  If CurrentQuery.FieldByName("SITUACAOLIVRO").Value <> "1" Then
    CanContinue = False
    SQLRotFin.Active = False
    Set SQLRotFin = Nothing
    bsShowMessage("A rotina já foi processada", "E")
    Exit Sub
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing
End Sub
 
