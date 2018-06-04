'HASH: EC9A98B017210B86B7BF2A6121D1BE32
'#Uses "*bsShowMessage"

Public Sub DIRETORIOERRO_OnBtnClick()
	Dim Interface As Object
	Set Interface = CreateBennerObject("BSTISS.Rotinas")

	If CurrentQuery.State <> 1 Or CurrentQuery.State <> 3 Then
		CurrentQuery.Edit
	End If

    CurrentQuery.FieldByName("DIRETORIOERRO").AsString = Interface.Diretorio(CurrentSystem)
    Set Interface = Nothing
End Sub

Public Sub DIRETORIORECEBIDO_OnBtnClick()
	Dim Interface As Object
	Set Interface = CreateBennerObject("BSTISS.Rotinas")

	If CurrentQuery.State <> 1 Or CurrentQuery.State <> 3 Then
		CurrentQuery.Edit
	End If

    CurrentQuery.FieldByName("DIRETORIORECEBIDO").AsString = Interface.Diretorio(CurrentSystem)
    Set Interface = Nothing
End Sub

Public Sub GRAUALUGUEIS_OnEnter()
	' SMS - 77458 - DRUMMOND - 19/04/2007
	Dim vsSQLEsp As String
	If Not CurrentQuery.FieldByName("EVENTOALUGUEIS").IsNull Then
		If CurrentQuery.FieldByName("EVENTOALUGUEIS").AsInteger > 0 Then
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp = vsSQLEsp + "             FROM SAM_TGE_GRAU B"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_GRAU     C ON (C.HANDLE = B.GRAU)"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_TIPOGRAU D ON (D.HANDLE = C.TIPOGRAU)"
			vsSQLEsp = vsSQLEsp + "            WHERE B.EVENTO = "+ CurrentQuery.FieldByName("EVENTOALUGUEIS").AsString +""
			vsSQLEsp = vsSQLEsp + "              AND D.CLASSIFICACAO = '0')"
			GRAUALUGUEIS.LocalWhere = vsSQLEsp
		End If
	Else
		GRAUALUGUEIS.LocalWhere = "1<>1"
	End If
End Sub

Public Sub GRAUDIARIA_OnEnter()
	' SMS - 77458 - DRUMMOND - 19/04/2007
	Dim vsSQLEsp As String
	If Not CurrentQuery.FieldByName("EVENTODIARIA").IsNull Then
		If CurrentQuery.FieldByName("EVENTODIARIA").AsInteger > 0 Then
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp = vsSQLEsp + "             FROM SAM_TGE_GRAU B"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_GRAU     C ON (C.HANDLE = B.GRAU)"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_TIPOGRAU D ON (D.HANDLE = C.TIPOGRAU)"
			vsSQLEsp = vsSQLEsp + "            WHERE B.EVENTO = "+ CurrentQuery.FieldByName("EVENTODIARIA").AsString +""
			vsSQLEsp = vsSQLEsp + "              AND D.CLASSIFICACAO = '3')"
			GRAUDIARIA.LocalWhere = vsSQLEsp
		End If
	Else
		GRAUDIARIA.LocalWhere = "1<>1"
	End If
End Sub

Public Sub GRAUGASES_OnEnter()
	' SMS - 77458 - DRUMMOND - 19/04/2007
	Dim vsSQLEsp As String
	If Not CurrentQuery.FieldByName("EVENTOGASES").IsNull Then
		If CurrentQuery.FieldByName("EVENTOGASES").AsInteger > 0 Then
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp = vsSQLEsp + "             FROM SAM_TGE_GRAU B"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_GRAU     C ON (C.HANDLE = B.GRAU)"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_TIPOGRAU D ON (D.HANDLE = C.TIPOGRAU)"
			vsSQLEsp = vsSQLEsp + "            WHERE B.EVENTO = "+ CurrentQuery.FieldByName("EVENTOGASES").AsString +""
			vsSQLEsp = vsSQLEsp + "              AND D.CLASSIFICACAO = '5')"
			GRAUGASES.LocalWhere = vsSQLEsp
		End If
	Else
		GRAUGASES.LocalWhere = "1<>1"
	End If
End Sub

Public Sub GRAUMATERIAL_OnEnter()
	' SMS - 77458 - DRUMMOND - 19/04/2007
	Dim vsSQLEsp As String
	If Not CurrentQuery.FieldByName("EVENTOMATERIAL").IsNull Then
		If CurrentQuery.FieldByName("EVENTOMATERIAL").AsInteger > 0 Then
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp = vsSQLEsp + "             FROM SAM_TGE_GRAU B"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_GRAU     C ON (C.HANDLE = B.GRAU)"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_TIPOGRAU D ON (D.HANDLE = C.TIPOGRAU)"
			vsSQLEsp = vsSQLEsp + "            WHERE B.EVENTO = "+ CurrentQuery.FieldByName("EVENTOMATERIAL").AsString +""
			vsSQLEsp = vsSQLEsp + "              AND D.CLASSIFICACAO = '1')"
			GRAUMATERIAL.LocalWhere = vsSQLEsp
		End If
	Else
		GRAUMATERIAL.LocalWhere = "1<>1"
	End If
End Sub

Public Sub GRAUMEDICAMENTO_OnEnter()
	' SMS - 77458 - DRUMMOND - 19/04/2007
	Dim vsSQLEsp As String
	If Not CurrentQuery.FieldByName("EVENTOMEDICAMENTO").IsNull Then
		If CurrentQuery.FieldByName("EVENTOMEDICAMENTO").AsInteger > 0 Then
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp = vsSQLEsp + "             FROM SAM_TGE_GRAU B"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_GRAU     C ON (C.HANDLE = B.GRAU)"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_TIPOGRAU D ON (D.HANDLE = C.TIPOGRAU)"
			vsSQLEsp = vsSQLEsp + "            WHERE B.EVENTO = "+ CurrentQuery.FieldByName("EVENTOMEDICAMENTO").AsString +""
			vsSQLEsp = vsSQLEsp + "              AND D.CLASSIFICACAO = '2')"
			GRAUMEDICAMENTO.LocalWhere = vsSQLEsp
		End If
	Else
		GRAUMEDICAMENTO.LocalWhere = "1<>1"
	End If
End Sub

Public Sub GRAUTAXA_OnEnter()
	' SMS - 77458 - DRUMMOND - 19/04/2007
	Dim vsSQLEsp As String
	If Not CurrentQuery.FieldByName("EVENTOTAXA").IsNull Then
		If CurrentQuery.FieldByName("EVENTOTAXA").AsInteger > 0 Then
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp =            "HANDLE IN (SELECT B.GRAU"
			vsSQLEsp = vsSQLEsp + "             FROM SAM_TGE_GRAU B"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_GRAU     C ON (C.HANDLE = B.GRAU)"
			vsSQLEsp = vsSQLEsp + "             JOIN SAM_TIPOGRAU D ON (D.HANDLE = C.TIPOGRAU)"
			vsSQLEsp = vsSQLEsp + "            WHERE B.EVENTO = "+ CurrentQuery.FieldByName("EVENTOTAXA").AsString +""
			vsSQLEsp = vsSQLEsp + "              AND D.CLASSIFICACAO = '4')"
			GRAUTAXA.LocalWhere = vsSQLEsp
		End If
	Else
		GRAUTAXA.LocalWhere = "1<>1"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	' SMS - 77458 - DRUMMOND - 19/04/2007 - INICIO
	If CurrentQuery.FieldByName("EVENTOALUGUEIS").AsInteger > 0 Then
		If Not CurrentQuery.FieldByName("GRAUALUGUEIS").AsInteger > 0 Then
			MsgBox "Por favor, selecione um grau para o evento 'Aluguéis'"
			CanContinue = False
		End If
	End If
	If CurrentQuery.FieldByName("EVENTODIARIA").AsInteger > 0 Then
		If Not CurrentQuery.FieldByName("GRAUDIARIA").AsInteger > 0 Then
			MsgBox "Por favor, selecione um grau para o evento 'Diária'"
			CanContinue = False
		End If
	End If
	If CurrentQuery.FieldByName("EVENTOGASES").AsInteger > 0 Then
		If Not CurrentQuery.FieldByName("GRAUGASES").AsInteger > 0 Then
			MsgBox "Por favor, selecione um grau para o evento 'Gases'"
			CanContinue = False
		End If
	End If
	If CurrentQuery.FieldByName("EVENTOMATERIAL").AsInteger > 0 Then
		If Not CurrentQuery.FieldByName("GRAUMATERIAL").AsInteger > 0 Then
			MsgBox "Por favor, selecione um grau para o evento 'Material'"
			CanContinue = False
		End If
	End If
	If CurrentQuery.FieldByName("EVENTOMEDICAMENTO").AsInteger > 0 Then
		If Not CurrentQuery.FieldByName("GRAUMEDICAMENTO").AsInteger > 0 Then
			MsgBox "Por favor, selecione um grau para o evento 'Medicamento'"
			CanContinue = False
		End If
	End If
	If CurrentQuery.FieldByName("EVENTOTAXA").AsInteger > 0 Then
		If Not CurrentQuery.FieldByName("GRAUTAXA").AsInteger > 0 Then
			MsgBox "Por favor, selecione um grau para o evento 'Taxa'"
			CanContinue = False
		End If
	End If
	' SMS - 77458 - DRUMMOND - 19/04/2007 - FIM
	If CurrentQuery.FieldByName("INTEGRACAOORIZON").AsInteger <> 2 Then
		CurrentQuery.FieldByName("CNPJORIZON").Value = Null
	End If

    While Mid(CurrentQuery.FieldByName("CAMINHOARQUIVOSAGENDADOS").AsString,Len(CurrentQuery.FieldByName("CAMINHOARQUIVOSAGENDADOS").AsString),1) = "\"
      CurrentQuery.FieldByName("CAMINHOARQUIVOSAGENDADOS").AsString = Left(CurrentQuery.FieldByName("CAMINHOARQUIVOSAGENDADOS").AsString,(Len(CurrentQuery.FieldByName("CAMINHOARQUIVOSAGENDADOS").AsString)-1))
    Wend

    While Mid(CurrentQuery.FieldByName("CAMINHOARQUIVOXMLLOTE").AsString,Len(CurrentQuery.FieldByName("CAMINHOARQUIVOXMLLOTE").AsString),1) = "\"
      CurrentQuery.FieldByName("CAMINHOARQUIVOXMLLOTE").AsString = Left(CurrentQuery.FieldByName("CAMINHOARQUIVOXMLLOTE").AsString,(Len(CurrentQuery.FieldByName("CAMINHOARQUIVOXMLLOTE").AsString)-1))
    Wend

	If UCase(CurrentQuery.FieldByName("CAMINHOARQUIVOSAGENDADOS").AsString) = UCase(CurrentQuery.FieldByName("CAMINHOARQUIVOXMLLOTE").AsString) Then
		bsShowMessage("A pasta de agendamento TISS e a pasta de Importação de XML em Lote devem ser diferentes.","I")
		CanContinue = False
	End If

End Sub
