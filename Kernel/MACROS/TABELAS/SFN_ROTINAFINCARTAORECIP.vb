'HASH: 9C9BDB37F59A8D8F2D814C128072B16E
'#Uses "*bsShowMessage"

Public Function VerificaSeProcessada(SITUACAO As String) As Boolean
  Dim SQLRotFin As Object

  Set SQLRotFin = NewQuery
  VerificaSeProcessada = True
  SQLRotFin.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQLRotFin.Active = True
  If (SITUACAO = "P") Then
    If SQLRotFin.FieldByName("SITUACAO").Value = "P" Then
      VerificaSeProcessada = False
      bsShowMessage("A Rotina já foi processada.", "I")
    End If
  Else
    If SQLRotFin.FieldByName("SITUACAO").Value = "A" Then
      VerificaSeProcessada = False
      bsShowMessage("A Rotina ainda não foi processada.", "I")
    End If
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing
End Function

Public Sub BOTAOCANCELAR_OnClick()
  	If CurrentQuery.State <> 1 Then
    	bsShowMessage("Os parâmetros não podem estar em edição", "I")
    	Exit Sub
	End If
	If (VerificaSeProcessada("C")) Then
		If VisibleMode Then
  			Set obj = CreateBennerObject("BSINTERFACE0061.RotinaFaturamentoCartaoRcp")
  			obj.CancelarFaturamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  		Else
			Dim vsmensagemerro As String
			Dim viRetorno As Long
			Dim vcContainer As CSDContainer
			Set vcContainer = NewContainer
			vcContainer.AddFields("HANDLE:INTEGER")
			vcContainer.Insert
			vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

			Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRetorno = obj.ExecucaoImediata(CurrentSystem, _
												"BSBEN016", _
												"CancelarFaturamentoCartaoRcp", _
												"Cancelamento das Faturas de Cartões de Reciprocidade", _
												CurrentQuery.FieldByName("HANDLE").AsInteger, _
												"SFN_ROTINAFINCARTAORECIP", _
												"SITUACAO", _
												"", _
												"", _
												"C", _
												False, _
												vsmensagemerro, _
												vcContainer)
	    	If viRetorno = 0 Then
    	  		bsShowMessage("Processo enviado para execução no servidor!", "I")
    		Else
      			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsmensagemerro, "I")
    		End If
    	End If
	End If
	Set obj = Nothing
 	''Cancelamento das Faturas de Cartões de Reciprocidade'
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  	If CurrentQuery.State <> 1 Then
    	bsShowMessage("Os parâmetros não podem estar em edição", "I")
    	Exit Sub
	End If

	Dim obj As Object
	If (VerificaSeProcessada("P")) Then
		If VisibleMode Then
  			Set obj = CreateBennerObject("BSINTERFACE0061.RotinaFaturamentoCartaoRcp")
  			obj.ProcessarFaturamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  		Else
			Dim vsmensagemerro As String
			Dim viRetorno As Long
			Dim vcContainer As CSDContainer
			Set vcContainer = NewContainer
			vcContainer.AddFields("HANDLE:INTEGER")
			vcContainer.Insert
			vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

			Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRetorno = obj.ExecucaoImediata(CurrentSystem, _
												"BSBEN016", _
												"ProcessarFaturamentoCartaoRcp", _
												"Faturamento de Cartões de Reciprocidade", _
												CurrentQuery.FieldByName("HANDLE").AsInteger, _
												"SFN_ROTINAFINCARTAORECIP", _
												"SITUACAO", _
												"", _
												"", _
												"P", _
												False, _
												vsmensagemerro, _
												vcContainer)
	    	If viRetorno = 0 Then
    	  		bsShowMessage("Processo enviado para execução no servidor!", "I")
    		Else
      			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsmensagemerro, "I")
    		End If
    	End If
	End If
	Set obj = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If (Not VerificaSeProcessada("P")) Then
    CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If (Not VerificaSeProcessada("P")) Then
    CanContinue = False
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
