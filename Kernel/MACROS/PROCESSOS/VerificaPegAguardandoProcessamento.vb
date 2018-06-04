'HASH: AB6D83FB09938ED9A2005BA036770F7A

Public Sub Main

	Dim qAjustaPegsProcessando As Object
	Set qAjustaPegsProcessando = NewQuery

	qAjustaPegsProcessando.Clear
	qAjustaPegsProcessando.Add("UPDATE SAM_PEG ")
	qAjustaPegsProcessando.Add("   SET SITUACAOPROCESSAMENTO = '1'")
	qAjustaPegsProcessando.Add(" WHERE SITUACAOPROCESSAMENTO <> '1' ")
	qAjustaPegsProcessando.Add("   AND DATA < :DATA")
	qAjustaPegsProcessando.ParamByName("DATA").AsDateTime = ServerNow - 0.05

	StartTransaction

	qAjustaPegsProcessando.ExecSQL

	Commit

	Set qAjustaPegsProcessando = Nothing
End Sub
