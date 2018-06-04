'HASH: 371F86D56B05C11BDB6CB9EE8D092E44
'Macro: SAM_MALADIRETA
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TIPOMODELOMALADIRETA_OnChange()
	If CurrentQuery.FieldByName("TIPOMODELOMALADIRETA").AsString <> "" Then
		Dim qMensagem As Object
		Set qMensagem = NewQuery

		qMensagem.Add("SELECT TEXTO FROM SAM_MENSAGEM_HTML WHERE HANDLE = :HSELECIONADO")
		qMensagem.ParamByName("HSELECIONADO").Value = CurrentQuery.FieldByName("TIPOMODELOMALADIRETA").AsInteger
		qMensagem.Active = True

		If qMensagem.FieldByName("TEXTO").AsString <> "" Then
			CurrentQuery.FieldByName("CORPO").AsString = qMensagem.FieldByName("TEXTO").AsString
		Else
			CurrentQuery.FieldByName("CORPO").Clear
		End If

		Set qMensagem = Nothing
	Else
		CurrentQuery.FieldByName("CORPO").Clear
	End If
End Sub

Public Sub EVENTOFINAL_OnExit()
	CheckEventos("F")
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)

	If CurrentQuery.FieldByName("MASCARATGE").AsString <> "" Then
		Dim vEventoFinal As Long

		vEventoFinal = AbreConsultaTGE(EVENTOFINAL.LocateText, ShowPopup)

		If vEventoFinal <> 0 Then
			CurrentQuery.FieldByName("EVENTOFINAL").AsInteger = vEventoFinal
		End If
	End If

End Sub

Public Sub EVENTOINICIAL_OnExit()
	CheckEventos("I")
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)

	If CurrentQuery.FieldByName("MASCARATGE").AsString <> "" Then
		Dim vEventoInicial As Long

		vEventoInicial = AbreConsultaTGE(EVENTOINICIAL.LocateText, ShowPopup)

		If vEventoInicial <> 0 Then
			CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = vEventoInicial
		End If
	End If

End Sub

Function AbreConsultaTGE(pCampo As String, ShowPopup As Boolean) As Long
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If (EVENTOINICIAL.PopupCase <> 0) Or (EVENTOFINAL.PopupCase <> 0) Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Estrutura Numérica|Estrutura TGE|Descrição TGE"
		vColunas = "SAM_TGE.ESTRUTURANUMERICA|SAM_TGE.ESTRUTURA|SAM_TGE.DESCRICAO"
		vColunas = vColunas
		vCriterio = "SAM_TGE.MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString + " AND ULTIMONIVEL = 'S' "
		vTabela = "SAM_TGE"
		vTitulo = "Tabela Geral de Eventos"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecs, vCriterio, vTitulo, True, pCampo)

		If vHandle <> 0 Then
			CurrentQuery.Edit
			AbreConsultaTGE = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If

End Function

Function CheckEventos(pCampo As String) As Boolean

	If Not CurrentQuery.FieldByName("EVENTOINICIAL").IsNull Then
		If CurrentQuery.FieldByName("EVENTOINICIAL").Value <> CurrentQuery.FieldByName("EVENTOFINAL").Value Then

			Dim qEventoInicial, qEventoFinal As Object
			Set qEventoInicial = NewQuery

			qEventoInicial.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :EVENTOINICIAL")

			qEventoInicial.ParamByName("EVENTOINICIAL").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
			qEventoInicial.Active = True

			Set qEventoFinal = NewQuery

			qEventoFinal.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :EVENTOFINAL")

			qEventoFinal.ParamByName("EVENTOFINAL").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
			qEventoFinal.Active = True

			If qEventoFinal.FieldByName("ESTRUTURA").Value < qEventoInicial.FieldByName("ESTRUTURA").Value Then
				bsShowMessage("Evento final não pode ser menor que o evento inicial!", "E")

				If pCampo = "I" Then
					CurrentQuery.FieldByName("EVENTOINICIAL").Clear
					EVENTOINICIAL.SetFocus
				Else
					CurrentQuery.FieldByName("EVENTOFINAL").Clear
					EVENTOFINAL.SetFocus
				End If

			End If

			Set qEventoInicial = Nothing
			Set qEventoFinal = Nothing
		End If
	End If

End Function
