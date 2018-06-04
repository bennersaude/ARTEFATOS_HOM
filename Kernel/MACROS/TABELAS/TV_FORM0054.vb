'HASH: 68C027C4780EAF6EF4DF192BD13A6C27
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub TABLE_AfterScroll()
       Dim vsAlertas As String
       Dim vsRecomendacoes As String  'Utilizado no desktop somente...


	   CriaTabelaTemporariaSQLServer
	   Set interface = CreateBennerObject("SAMPEG.Rotinas")
	   vsAlertas = interface.VerAlertas(CurrentSystem, RecordHandleOfTable("SAM_GUIA_EVENTOS"), vsRecomendacoes)

	   Set interface = Nothing
       If vsAlertas <> "" Then
		   ROTALERTAS.Text = "@" + vsAlertas
	   Else
	   	   ROTALERTAS.Text = "Não há alertas a serem mostrados"
	   End If
End Sub
