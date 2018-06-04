'HASH: 54E35C4D78BF241433B22D7711585116
'Macro: SFN_ROTINAFIN
'A funcao NodeInternalCode é utilizada para determinar se a carga correspondente é da Tarefas de Modelo,
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina
'Alteração: 26/12/2005
'      SMS: 52120 - Marcelo Barbosa
'#Uses "*bsShowMessage"

Public Sub BOTAOAGENDAR_OnClick()
  Dim qr As Object
  Dim qr1 As Object
  Dim vSituacao As String
  Dim vTabela As String
  Dim vLegendaAgendamento As String
  Dim VLegendaAberta As String
  Dim VLegendaProcessada As String
  Set qr = NewQuery
  Set qr1 = NewQuery

  vTabela = "SFN_ROTINAFIN"
  vLegendaAgendamento = "G"
  VLegendaAberta = "A"
  VLegendaProcessada = "P"

  qr.Clear

  qr.Add("SELECT SITUACAO FROM " + vTabela + " WHERE HANDLE = :pHANDLE")

  qr.ParamByName("pHandle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qr.Active = True

  vSituacao = qr.FieldByName("SITUACAO").AsString

  If vSituacao <> vLegendaAgendamento Then
	If vSituacao = VLegendaAberta Then
	  If bsShowMessage("Confirme o agendamento da rotina", "Q") = vbYes Then '(6=yes, 7=não)
		qr1.Clear

      If Not InTransaction Then StartTransaction
		qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")

		qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		qr1.ParamByName("pSituacao").AsString = vLegendaAgendamento

		qr1.ExecSQL
	  If InTransaction Then Commit
	  End If
	Else
	  bsShowMessage("Rotina já foi processada.", "I")
	End If
  Else
	If bsShowMessage("O agendamento da rotina será retirado. Confirme para continuar.", "Q") = vbYes Then
	  qr1.Clear
  	 	If Not InTransaction Then StartTransaction
		  	qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
		  	qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  		If (CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull) Then
				qr1.ParamByName("pSituacao").AsString = VLegendaAberta
		  	Else
				qr1.ParamByName("pSituacao").AsString = VLegendaProcessada
	  		End If
	  		qr1.ExecSQL
	  	If InTransaction Then Commit
	End If
  End If

  Set qr = Nothing
  Set qr1 = Nothing

  If VisibleMode Then
	SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If VisibleMode Then
    If Not CurrentQuery.FieldByName("TIPOFATURAMENTO").IsNull Then
      Dim sql As Object
      Set sql = NewQuery

      sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = :HANDLE")

      sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
      sql.Active = True

      If sql.FieldByName("CODIGO").AsInteger = 210 Or _
        sql.FieldByName("CODIGO").AsInteger = 310 Or _
        sql.FieldByName("CODIGO").AsInteger = 110 Or _
        sql.FieldByName("CODIGO").AsInteger = 120 Or _
        sql.FieldByName("CODIGO").AsInteger = 130 Or _
        sql.FieldByName("CODIGO").AsInteger = 140 Then
          SITUACAO.Visible = False
        Else
          SITUACAO.Visible = True
      End If
    End If
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If NodeInternalCode <> 800 Then
    Dim qSQLRotina As Object
    Set qSQLRotina = NewQuery

    qSQLRotina.Clear
    qSQLRotina.Add("SELECT SFAT.CODIGO")
    qSQLRotina.Add("FROM SIS_TIPOFATURAMENTO SFAT")
    qSQLRotina.Add("WHERE SFAT.HANDLE = :HTIPOFATURAMENTO")
    qSQLRotina.ParamByName("HTIPOFATURAMENTO").Value = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
    qSQLRotina.Active = True

    If (qSQLRotina.FieldByName("CODIGO").AsInteger = 110) Or _
       (qSQLRotina.FieldByName("CODIGO").AsInteger = 120) Or _
       (qSQLRotina.FieldByName("CODIGO").AsInteger = 130) Or _
       (qSQLRotina.FieldByName("CODIGO").AsInteger = 140) Then
      qSQLRotina.Clear
      qSQLRotina.Add("SELECT RFAT.HANDLE, RFAT.SITUACAOFATURAMENTO, RFAT.SITUACAOPARCELAMENTO")
      qSQLRotina.Add("FROM SFN_ROTINAFINFAT RFAT")
      qSQLRotina.Add("WHERE RFAT.ROTINAFIN = :HROTINAFIN")
      qSQLRotina.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      qSQLRotina.Active = True
      If qSQLRotina.FieldByName("HANDLE").AsInteger > 0 Then
        If Not(qSQLRotina.FieldByName("SITUACAOFATURAMENTO").AsString  = "1") Or _
           Not(qSQLRotina.FieldByName("SITUACAOPARCELAMENTO").AsString = "1") Then
          Set qSQLRotina = Nothing
	      CanContinue = False
	      bsShowMessage("A Rotina não está aberta!", "E")
	      Exit Sub
	    End If
      End If
    Else
	  If CurrentQuery.FieldByName("SITUACAO").Value = "P" Then
	    CanContinue = False
	    bsShowMessage("A Rotina já foi processada!", "E")
	    Exit Sub
	  End If
	End If

	Set qSQLRotina = Nothing
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If NodeInternalCode <> 800 Then
    Dim qSQLRotina As Object
    Set qSQLRotina = NewQuery

    qSQLRotina.Clear
    qSQLRotina.Add("SELECT SFAT.CODIGO")
    qSQLRotina.Add("FROM SIS_TIPOFATURAMENTO SFAT")
    qSQLRotina.Add("WHERE SFAT.HANDLE = :HTIPOFATURAMENTO")
    qSQLRotina.ParamByName("HTIPOFATURAMENTO").Value = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
    qSQLRotina.Active = True

    If (qSQLRotina.FieldByName("CODIGO").AsInteger = 110) Or _
       (qSQLRotina.FieldByName("CODIGO").AsInteger = 120) Or _
       (qSQLRotina.FieldByName("CODIGO").AsInteger = 130) Or _
       (qSQLRotina.FieldByName("CODIGO").AsInteger = 140) Then
      qSQLRotina.Clear
      qSQLRotina.Add("SELECT RFAT.HANDLE, RFAT.SITUACAOFATURAMENTO, RFAT.SITUACAOPARCELAMENTO")
      qSQLRotina.Add("FROM SFN_ROTINAFINFAT RFAT")
      qSQLRotina.Add("WHERE RFAT.ROTINAFIN = :HROTINAFIN")
      qSQLRotina.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      qSQLRotina.Active = True
      If qSQLRotina.FieldByName("HANDLE").AsInteger > 0 Then
        If Not(qSQLRotina.FieldByName("SITUACAOFATURAMENTO").AsString  = "1") Or _
           Not(qSQLRotina.FieldByName("SITUACAOPARCELAMENTO").AsString = "1") Then
          Set qSQLRotina = Nothing
	      CanContinue = False
	      bsShowMessage("A Rotina não está aberta!", "E")
	      Exit Sub
	    End If
        bsShowMessage("Excluir os itens da pasta Parâmetros e Contratos, encontrados abaixo!", "E")
        CanContinue = False
        Exit Sub
      End If
    Else
	  If CurrentQuery.FieldByName("SITUACAO").Value = "P" Then
	    CanContinue = False
	    bsShowMessage("A Rotina já foi processada!", "E")
	    Exit Sub
	  End If
	End If

	Set qSQLRotina = Nothing
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
    If CurrentQuery.State = 2 Then
      Dim qSQLRotina As Object
      Set qSQLRotina = NewQuery

      qSQLRotina.Clear
      qSQLRotina.Add("SELECT SFAT.CODIGO")
      qSQLRotina.Add("FROM SIS_TIPOFATURAMENTO SFAT")
      qSQLRotina.Add("WHERE SFAT.HANDLE = :HTIPOFATURAMENTO")
      qSQLRotina.ParamByName("HTIPOFATURAMENTO").Value = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
      qSQLRotina.Active = True

      If (qSQLRotina.FieldByName("CODIGO").AsInteger = 110) Or _
         (qSQLRotina.FieldByName("CODIGO").AsInteger = 120) Or _
         (qSQLRotina.FieldByName("CODIGO").AsInteger = 130) Or _
         (qSQLRotina.FieldByName("CODIGO").AsInteger = 140) Then
        qSQLRotina.Clear
        qSQLRotina.Add("SELECT RFAT.HANDLE, RFAT.SITUACAOFATURAMENTO, RFAT.SITUACAOPARCELAMENTO")
        qSQLRotina.Add("FROM SFN_ROTINAFINFAT RFAT")
        qSQLRotina.Add("WHERE RFAT.ROTINAFIN = :HROTINAFIN")
        qSQLRotina.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
        qSQLRotina.Active = True
        If qSQLRotina.FieldByName("HANDLE").AsInteger > 0 Then
          If Not(qSQLRotina.FieldByName("SITUACAOFATURAMENTO").AsString  = "1") Or _
             Not(qSQLRotina.FieldByName("SITUACAOPARCELAMENTO").AsString = "1") Then
            Set qSQLRotina = Nothing
	        CanContinue = False
	        bsShowMessage("A Rotina já foi processada!", "E")
	        Exit Sub
	      End If
        End If
      Else
	    If CurrentQuery.FieldByName("SITUACAO").Value = "P" Then
	      CanContinue = False
	      bsShowMessage("A Rotina já foi processada!", "E")
	      Exit Sub
	    End If
	  End If

	  Set qSQLRotina = Nothing
	End If

	Dim SQLPARAMFIN As Object
	Set SQLPARAMFIN = NewQuery

	SQLPARAMFIN.Clear

	SQLPARAMFIN.Active = False

	SQLPARAMFIN.Add("SELECT PERMITEDATACONTABILMAIOR FROM SFN_PARAMETROSFIN")

	SQLPARAMFIN.Active = True

	If (CurrentQuery.FieldByName("TABDATACONTABIL").AsInteger = 1) Then
	  If (CurrentQuery.FieldByName("DATAROTINA").AsDateTime > CurrentQuery.FieldByName("DATACONTABIL").AsDateTime) Then
		bsShowMessage("Data contábil menor que a data da rotina. Verifique!", "E")
	  End If

      'Faturamento de INSS
	  If (VisibleMode And _
	      (NodeInternalCode <> 610)) Or _
	     (WebMode And _
	      (WebMenuCode = "T2392")) Then
		'  Verificar se a data contábil está dentro da competência
		Dim SQLCompetFin As Object
		Set SQLCompetFin = NewQuery

		SQLCompetFin.Add("SELECT COMPETENCIA FROM SFN_COMPETFIN")
		SQLCompetFin.Add("WHERE HANDLE = :HANDLECOMPETFIN")

		SQLCompetFin.ParamByName("HANDLECOMPETFIN").Value = CurrentQuery.FieldByName("COMPETFIN").Value
		SQLCompetFin.Active = True

		If ((Month(CurrentQuery.FieldByName("DATACONTABIL").Value) <> Month(SQLCompetFin.FieldByName("COMPETENCIA").Value)) Or _
			(Year(CurrentQuery.FieldByName("DATACONTABIL").Value) <> Year(SQLCompetFin.FieldByName("COMPETENCIA").Value))) And _
		   (SQLPARAMFIN.FieldByName("PERMITEDATACONTABILMAIOR").AsString = "N") Then
		  CanContinue = False
		  bsShowMessage("A data contábil deve estar dentro da competência da rotina financeira", "E")
		  Exit Sub
		End If

		SQLCompetFin.Active = False

		Set SQLCompetFin=Nothing
	  End If
	End If

	If (Not CurrentQuery.FieldByName("TAREFA").IsNull) And (CurrentQuery.FieldByName("ORDEM").IsNull) Then
	  bsShowMessage("O campo 'Ordem' é obrigatório quando a rotina está associada a uma tarefa.", "E")
	  CanContinue = False
	End If

	If CurrentQuery.FieldByName("TABDATACONTABIL").Value="2" Then 'assumir vencimento da fatura
	  CurrentQuery.FieldByName("DATACONTABIL").Clear
	End If

    If CurrentQuery.State = 3 Then
	  Dim SEQUENCIA As Long

  	  NewCounter("SFN_ROTINAFIN", CurrentQuery.FieldByName("COMPETFIN").AsInteger, 1, SEQUENCIA)

	  CurrentQuery.FieldByName("SEQUENCIA").Value = SEQUENCIA
    End If

	If WebMode Then
      If (WebMenuCode = "T5955") Or (WebMenuCode = "T5959" )Then 'Luciano T. Alberti - 04/09/2008
        CurrentQuery.FieldByName("EHREAPRESENTACAO").AsString = "S"
      ElseIf (WebMenuCode = "T5956") Or (WebMenuCode = "T5960") Then
        CurrentQuery.FieldByName("CONTROLEPAGAMENTO").AsString = "S"
      End If
	Else
      If NodeInternalCode = 96895 Then 'Coelho SMS: 96895
        CurrentQuery.FieldByName("EHREAPRESENTACAO").AsString = "S"
      ElseIf NodeInternalCode = 96896 Then
        CurrentQuery.FieldByName("CONTROLEPAGAMENTO").AsString = "S"
      End If
    End If

End Sub

Public Sub TABLE_NewRecord()
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT PADRAODATACONTABIL FROM SFN_PARAMETROSFIN")

  sql.Active = True

  If sql.FieldByName("PADRAODATACONTABIL").IsNull Then
	bsShowMessage("Falta informação nos parâmetros financeiros (PADRAODATACONTABIL)", "I")

	sql.Active = False

	Set sql = Nothing

	Exit Sub
  End If

  'SMS 52120 - Marcelo Barbosa - 26/12/2005
  'Visualização do tab (Rotina ou Modelo) conforme a carga em que se encontra
  If NodeInternalCode = 800 Then
    If VisibleMode Then
	  'TABTIPOROTINA.Pages(0).Visible = False
	  'TABTIPOROTINA.Pages(1).Visible = True
	End If

	'CurrentQuery.FieldByName("TABTIPOROTINA").Value = 2
	CurrentQuery.FieldByName("TABDATACONTABILMODELO").Value = sql.FieldByName("PADRAODATACONTABIL").Value
  Else
    If VisibleMode Then
	  'TABTIPOROTINA.Pages(0).Visible = True
	  'TABTIPOROTINA.Pages(1).Visible = False
	End If

	'CurrentQuery.FieldByName("TABTIPOROTINA").Value = 1
	CurrentQuery.FieldByName("TABDATACONTABIL").Value = sql.FieldByName("PADRAODATACONTABIL").Value
  End If

  sql.Active = False

  Set sql = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
	Case "BOTAOAGENDAR"
	  BOTAOAGENDAR_OnClick
  End Select
End Sub
