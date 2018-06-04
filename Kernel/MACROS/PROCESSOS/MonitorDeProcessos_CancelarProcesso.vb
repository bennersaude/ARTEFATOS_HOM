'HASH: 278FE697016326E77707CB103ABB78C9

Public Sub Main
	Dim piHandleProc As Long
	Dim vsUsuario As String
	Dim vdData As Date

	piHandleProc = CLng( ServiceVar("piHandleProc") )

	Dim ServerExec As CSServerExec
	Set ServerExec = GetServerExec(piHandleProc)
	ServerExec.RequestAbort

	Set ServerExec = Nothing

	vdData = ServerNow

	Dim SQL As Object
	Set SQL = NewQuery
	SQL.Add("SELECT APELIDO FROM Z_GRUPOUSUARIOS WHERE HANDLE = " + CStr(CurrentUser))
	SQL.Active = True

	vsUsuario = SQL.FieldByName("APELIDO").AsString

	ServiceVar("psUsuarioAbort") = vsUsuario
	ServiceVar("psDataAbort") = CStr(vData)
End Sub
