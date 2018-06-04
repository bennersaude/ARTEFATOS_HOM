'HASH: DF3DA6F0FD536F6B944F1029F0A2C6D5
'#Uses "*addXMLAtt"

Public Sub Main
	Dim handleEvento As Long

	handleEvento = ServiceVar("psEvento")

	Dim sql As BPesquisa
	Set sql=NewQuery

	Dim xml As String

	On Error GoTo erro

	sql.Clear
	sql.Add("SELECT DISTINCT G.HANDLE, G.DESCRICAO, G.GRAU CODIGO")
	sql.Add("      FROM SAM_GRAU G										 ")
	sql.Add("     WHERE (( G.VERIFICAGRAUSVALIDOS = 'N'					 ")
	sql.Add("              OR (EXISTS (SELECT GE.HANDLE                  ")
	sql.Add("                                FROM SAM_TGE_GRAU GE        ")
	sql.Add("								WHERE GE.EVENTO=:HANDLE      ")
	sql.Add("									AND GE.GRAU=G.HANDLE)))) ")

	sql.ParamByName("HANDLE").AsInteger = handleEvento
	sql.Active=True

	xml="<registros>"
	While Not sql.EOF
		xml=xml + "<registro>"
		xml=xml + addXMLAtt( "handle", "handle", sql, "")
		xml=xml + addXMLAtt( "descricao", "descricao", sql, "caption='Descrição' width='300'")
		xml=xml + addXMLAtt( "codigo", "codigo", sql, "caption='Código' width='120'")
		xml=xml + "</registro>"
		sql.Next
	Wend
	xml=xml + "</registros>"

	ServiceVar("psResult") = CStr(xml)
	GoTo fim

	erro:

		ServiceVar("psResult") = CStr(Err.Description)

	fim:

	Set sql=Nothing

End Sub
