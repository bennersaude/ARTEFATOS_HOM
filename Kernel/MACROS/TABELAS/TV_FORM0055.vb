'HASH: B1D8BDD1A30130D6DDEBFFC7FCB1F2D8
 
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub TABLE_AfterScroll()
       Dim vsAlertas As String
       Dim vsRecomendacoes As String  'Utilizado no desktop somente...


	   CriaTabelaTemporariaSQLServer
	   Set interface = CreateBennerObject("SAMPEG.Rotinas")
	   vsAlertas = interface.VerAlertasPeg(CurrentSystem, RecordHandleOfTable("SAM_PEG"))

	   Set interface = Nothing

       If vsAlertas <> "" Then
	   	ROTALERTASPEG.Text = "@" + vsAlertas
	   Else
	   	ROTALERTASPEG.Text = "Não há alertas a serem mostrados"
	   End If
End Sub
