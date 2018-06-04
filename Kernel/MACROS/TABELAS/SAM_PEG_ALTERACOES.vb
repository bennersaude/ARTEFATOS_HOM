'HASH: 2F52380260D8CE41434B9C02BCA0F6F9
Public Sub TABLE_AfterScroll()
	DATACONTABILANTERIOR.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 1
	NOVADATACONTABIL.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 1

	DATAPAGAMENTOANTERIOR.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 2
	NOVADATAPAGAMENTO.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 2

	QTDGUIASAPRESENTADASANTERIOR.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 3
	NOVAQTDGUIASAPRESENTADAS.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 3

	VALORAPRESENTADOANTERIOR.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 4
	NOVOVALORAPRESENTADO.Visible = CurrentQuery.FieldByName("TABTIPO").AsInteger = 4
End Sub
