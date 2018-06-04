'HASH: D950EB3B013593ABD0E0EDDFC4B7452B
'#Uses "*bsShowMessage"

Option Explicit



Public Sub BOTAOCANCELAR_OnClick()

	Dim Interface As Object
    Dim vDescricaoRotina As String


	If CurrentQuery.State <>1 Then
	bsShowMessage("O Registro não podem estar em edição","I")
	Exit Sub
	End If

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT SITUACAO")
	SQL.Add("FROM SFN_ROTINAFIN")
	SQL.Add("WHERE HANDLE = :HROTINAFIN")
	SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
	SQL.Active = True

	If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 1 Then
		vDescricaoRotina = "Rotina de Cancelamento de Financiamento de PF"
	ElseIf CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 2 Then
		vDescricaoRotina = "Rotina de Cancelamento Exportação Folha"
	ElseIf CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 3 Then
		vDescricaoRotina = "Rotina de Cancelamento Importação Folha"
	End If

	UserVar("Descricao_Rotina_Cancelamento") = vDescricaoRotina


	If VisibleMode Then

		Dim vContainer As CSDContainer
		Set vContainer = NewContainer

		vContainer.AddFields("HANDLE:INTEGER;INTERFACE:STRING")
		vContainer.Insert
		vContainer.Field("HANDLE").AsInteger    = CurrentQuery.FieldByName("HANDLE").AsInteger
		vContainer.Field("INTERFACE").AsString  = "BSFin001.Financiamento_Cancelar"

	    Set Interface = CreateBennerObject("BSINTERFACE.Rotinas")
	    Interface.Executar(CurrentSystem, vContainer)
	    Set Interface = Nothing
	    Set vContainer = Nothing
	    Set SQL = Nothing
	Else

	    Dim vsMensagemErro As String
		Dim viRetorno As Long


		Dim obj As Object
	    Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	    viRetorno = obj.ExecucaoImediata(CurrentSystem, _
	                                 "BSFIN001", _
	                                 "Financiamento_Cancelar", _
	                                 vDescricaoRotina & " - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
	                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
	                                 "SFN_ROTINAFINPF", _
	                                 "SITUACAO", _
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

	  	Set obj = Nothing

	End If


  Set SQL = Nothing

  RefreshNodesWithTable("SFN_ROTINAFINPF")

End Sub

Public Sub BOTAOPROCESSAR_OnClick()
	Dim Interface As Object
	Dim vDescricaoRotina As String

	If CurrentQuery.State <>1 Then
		bsShowMessage("O Registro não podem estar em edição","I")
		Exit Sub
	End If

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT SITUACAO")
	SQL.Add("FROM SFN_ROTINAFIN")
	SQL.Add("WHERE HANDLE = :HROTINAFIN")
	SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
	SQL.Active = True

	If SQL.FieldByName("SITUACAO").AsString <>"A" Then
		bsShowMessage("A rotina já foi processada","I")
		Set SQL = Nothing
		Exit Sub
	End If

	SQL.Clear

	If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 1 Then 'processar
		SQL.Add("SELECT * FROM SFN_ROTINAFINPF_PARAM WHERE ROTINAFINPF =:ROTINA")
		SQL.ParamByName("ROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsString
		SQL.Active = True

		If SQL.EOF Then
		  bsShowMessage("Preencha os parâmetros de contratos a faturar","I")
		  Set SQL = Nothing
		  Exit Sub
		End If
	End If

	If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 1 Then
		vDescricaoRotina = "Rotina de Processamento de Financiamento de PF"
	ElseIf CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 2 Then
		vDescricaoRotina = "Rotina de Processamento Exportação Folha"
	ElseIf CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 3 Then
		vDescricaoRotina = "Rotina de Processamento Importação Folha"
	End If

	UserVar("Descricao_Rotina_Processamento") = vDescricaoRotina


	If VisibleMode Then

		Dim vContainer As CSDContainer
		Set vContainer = NewContainer

		vContainer.AddFields("HANDLE:INTEGER;INTERFACE:STRING")
		vContainer.Insert
		vContainer.Field("HANDLE").AsInteger    = CurrentQuery.FieldByName("HANDLE").AsInteger
		vContainer.Field("INTERFACE").AsString  = "BSFin001.Financiamento_Processar"

	    Set Interface = CreateBennerObject("BSINTERFACE.Rotinas")
	    Interface.Executar(CurrentSystem, vContainer)
	    Set Interface = Nothing
	    Set vContainer = Nothing
	    Set SQL = Nothing
	Else

	    Dim vsMensagemErro As String
		Dim viRetorno As Long


		Dim obj As Object
	    Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	    viRetorno = obj.ExecucaoImediata(CurrentSystem, _
	                                 "BSFIN001", _
	                                 "Financiamento_Processar", _
	                                 vDescricaoRotina &  " - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
	                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
	                                 "SFN_ROTINAFINPF", _
	                                 "SITUACAO", _
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

	  	Set obj = Nothing

	End If

  RefreshNodesWithTable("SFN_ROTINAFINPF")

End Sub

Public Sub TABLE_AfterScroll()

  Dim qRotinaFin As Object
  Set qRotinaFin = NewQuery

  With qRotinaFin
    .Active = False
    .Clear
    .Add("SELECT SITUACAO")
    .Add("  FROM SFN_ROTINAFIN")
    .Add(" WHERE HANDLE = :HANDLE")
    .ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    .Active = True
  End With

  If (qRotinaFin.FieldByName("SITUACAO").AsString = "P") Or (qRotinaFin.FieldByName("SITUACAO").AsString = "S") Then
    BOTAOPROCESSAR.Enabled = False
    BOTAOCANCELAR.Enabled = True
  Else
    BOTAOPROCESSAR.Enabled = True
    BOTAOCANCELAR.Enabled = False
  End If

  Set qRotinaFin = Nothing

End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsInteger <> 1 Then
  		bsShowMessage("Alteração não permitida. Rotina já processada","E")
  		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 3 Then
  	If CurrentQuery.FieldByName("ARQUIVO").IsNull Then
  		bsShowMessage("Informe o arquivo para importação.","E")
  		CanContinue = False
  	End If
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "PROCESSAR
			BOTAOPROCESSAR_OnClick
		Case "CANCELAR"
			BOTAOCANCELAR_OnClick
	End Select
End Sub

Public Sub TABTIPOROTINA_OnChange()
  CurrentQuery.UpdateRecord
  If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger <> 1 Then
    ARQUIVO.ReadOnly = CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Or CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 2
    ARQUIVOCANCELADOS.ReadOnly = CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 2
  End If
End Sub
