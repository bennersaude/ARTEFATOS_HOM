'HASH: 57D7220C9265466417AA33398AA01EEE

Public Sub Main

	Dim vlHTabela As Long
	Dim vsTabela As String
	Dim vsAlterarCampos As String

	Set SQL = NewQuery

	vlHTabela = CLng(ServiceVar("HANDLETABELACLONE"))
	vsTabela = ServiceVar("TABELACLONE")
	vsAlterarCampos = ServiceVar("ALTERARCAMPOS")


	ServiceResult = Str(CopyRecord(vsTabela,vlHTabela,vsAlterarCampos))

End Sub
