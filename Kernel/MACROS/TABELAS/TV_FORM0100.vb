'HASH: 0B5172F2E798B54A471AEE999C83E720
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim msgArray() As String
	Dim valArray() As String
	Set Interface = CreateBennerObject("SamFaturamento.Faturamento")
  	msgSaida = Interface.AlterarDiaCobrancaWeb(CurrentQuery.FieldByName("HANDLECONTRATO").AsInteger, _
  	CurrentQuery.FieldByName("NOVODIACOBRANCA").AsInteger, _
  	CurrentQuery.FieldByName("TIPORECEBIMENTO").AsInteger)

	msgArray() = Split(msgSaida,";")


	msg = ""
	logSaida = "

	For i = 0 To UBound(msgArray)
		valArray() = Split(msgArray(i),":")
		If valArray(0) = "msg" Then
			msg = valArray(1)
		Else
			logSaida = logSaida + valArray(1)
		End If
	Next
	CurrentQuery.FieldByName("LOG").Value = logSaida
	bsShowMessage(msg, "I")
  	Set Interface = Nothing
End Sub
