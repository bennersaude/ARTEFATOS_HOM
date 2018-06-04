'HASH: 25D82A71EA7F73EADB662B98AF6BA0F4

Public Sub Main
	Dim HandleDoc As Long
	Dim Campo As String
	Dim Obj As Object
	Dim result As String
	Dim SQL As Object

	Set SQL = NewQuery

	HandleDoc = CLng(ServiceVar("DOCUMENTO"))
	Campo = ServiceVar("CAMPO")
	Set Obj=CreateBennerObject("SamImpressao.Boleto")

	SQL.Clear
	SQL.Add("SELECT AGENCIACODCEDENTE, CIP FROM SAM_PARAMETROSWEB")
	SQL.Active = True

	Select Case Campo

	  Case "AGENCIA"
	    result =  Obj.ImprimirBoletoWeb(CurrentSystem,HandleDoc,SQL.FieldByName("AGENCIACODCEDENTE").AsString)

	  Case "CIP"
	    result =  Obj.ImprimirBoletoWeb(CurrentSystem,HandleDoc,SQL.FieldByName("CIP").AsString)

	End Select

	Set Obj = Nothing
	SQL.Active = False
	Set SQL = Nothing
	ServiceResult = result

End Sub
