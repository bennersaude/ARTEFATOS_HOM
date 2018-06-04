'HASH: 59E16377055560B18A11CAD64BAFF894

Public Sub Main
	Dim Obj As Object
	Dim result As String
	Dim DataBase As Date
	Dim DataAtualizacao As Date
	Dim ValorBase As Currency
	Dim RegraFin As Integer
	Dim Natureza As String
	Dim Municipio As Integer
	Dim ValorJuro As Double
	Dim ValorMulta As Double
	Dim ValorCorrecao As Double
	Dim ValorDesconto As Double

	DataBase = CDate(ServiceVar("DATABASE"))
	DataAtualizacao = CDate(ServiceVar("DATAATUALIZACAO"))
	ValorBase = CCur(ServiceVar("VALORBASE"))
	RegraFin = CInt(ServiceVar("REGRAFINANCEIRA"))
	Natureza = CStr(ServiceVar("NATUREZA"))
	Municipio = CInt(ServiceVar("MUNICIPIO"))


	Set Obj=CreateBennerObject("Financeiro.Geral")
	Obj.Financeira(CurrentSystem, RegraFin, DataBase, DataAtualizacao, ValorBase, Natureza, Municipio, ValorJuro, ValorMulta, ValorCorrecao, ValorDesconto)
	Set Obj = Nothing

	ServiceResult = CStr(ValorMulta)+"|"+CStr(ValorJuro)

End Sub
