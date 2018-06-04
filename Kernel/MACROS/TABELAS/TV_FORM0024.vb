'HASH: 88276201E0D698F55FB456EF78D3B4F1
'#uses "*CriaTabelaTemporariaSqlServer"
Option Explicit

Public Function PegarAlertas As String

	Dim Mensagem As String
	Dim alertas As String
	Dim resultado As Integer
	Dim evento As Long

	Dim dll As Object

	Set dll=CreateBennerObject("samauto.autorizador")

	evento = RecordHandleOfTable("SAM_AUTORIZ_EVENTOGERADO")
	If evento > 0 Then
		resultado = dll.PegarAlertasGerado(CurrentSystem, evento, False, False, alertas, Mensagem)
	Else
		resultado = dll.PegarAlertasSolicitado(CurrentSystem, RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT"), False, False, "A", alertas, Mensagem)
	End If
	Set dll=Nothing
	If resultado > 0 Then
		InfoDescription = Mensagem
		PegarAlertas= ""
	Else
		PegarAlertas=alertas
	End If

End Function

Public Function montaAlertas (alertas As String)
	Dim Mensagem As String
	Dim alertasFormatados As String
	Dim resultado As Integer

	Dim dll As Object
	Set dll = CreateBennerObject("SAMAUTO.AUTORIZADOR")
	resultado = dll.MostrarAlertas(CurrentSystem, alertas, alertasFormatados, Mensagem)
	Set dll=Nothing
	If resultado > 0 Then
		InfoDescription = Mensagem
		montaAlertas= ""
	Else
		montaAlertas=alertasFormatados
	End If
End Function

Public Sub TABLE_AfterScroll()
'carregar o rótulo com os alertas no formato HTML
	ROTULOALERTA.Text = ""

    CriaTabelaTemporariaSqlServer

	Dim alertas As String
	alertas = PegarAlertas
	If alertas <> "" Then
	    Dim Mensagem As String
	    Mensagem = montaAlertas(alertas)
		ROTULOALERTA.Text = "@" + Replace(Mensagem, "&nbsp", " ", 1, Len(Mensagem))
	End If
End Sub
