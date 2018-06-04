'HASH: 7CA1B4E339A6654BD0CD26FC533F1B40
'#Uses "*addXML"

Sub Main
	On Error GoTo erro

	Dim sql As BPesquisa

	Dim psCampos As String
	Dim psFrom As String
	Dim psWhere As String
	Dim psOrderby As String
	Dim piRegistros As Long
	Dim psMensagem As String

	Dim vsMarcadorTpData_Inicio As String
	Dim vsMarcadorTpData_Fim As String
	Dim viPosIniCpoTpData As Integer
	Dim viPosFimCpoTpData As Integer
	Dim vsCampoTipoData As String

	psCampos = CStr( ServiceVar("psCampos") )
	psFrom = CStr( ServiceVar("psFrom") )
	psWhere = CStr( ServiceVar("psWhere") )
	psOrderby = CStr( ServiceVar("psOrderby") )
	piRegistros = CLng( ServiceVar("piRegistros") )
	psMensagem = CStr( ServiceVar("psMensagem") )

	'Luciano T. Alberti - SMS 90450 - 09/04/2008 - Início
	' Tratamento de campo tipo data
	vsMarcadorTpData_Inicio = "#CPOTPDATA"
	vsMarcadorTpData_Fim = "CPOTPDATA#"
	vsCampoTipoData = ""

	If InStr(psWhere, vsMarcadorTpData_Inicio) Then
		viPosIniCpoTpData = StrPos(vsMarcadorTpData_Inicio, psWhere)
		viPosFimCpoTpData = StrPos(vsMarcadorTpData_Fim, psWhere)
		vsCampoTipoData = Mid(psWhere, viPosIniCpoTpData, viPosFimCpoTpData - viPosIniCpoTpData + Len(vsMarcadorTpData_Fim))
		
		psWhere = Replace(psWhere, vsCampoTipoData, ":CPOTPDATA", 1, Len(psWhere))
		vsCampoTipoData = Replace(vsCampoTipoData, vsMarcadorTpData_Inicio, "", 1, Len(vsCampoTipoData))
		vsCampoTipoData = Replace(vsCampoTipoData, vsMarcadorTpData_Fim, "", 1, Len(vsCampoTipoData))
	End If
	'Luciano T. Alberti - SMS 90450 - 09/04/2008 - Início

	Set sql = NewQuery

	sql.Add( "SELECT COUNT(1) REG")
	sql.Add( "    FROM " + psFrom )

	If Not ( psWhere = "" ) Then
		psWhere = Replace( Replace( psWhere, "&lt", ">" ), "&gt", "<")
		sql.Add( "WHERE " + psWhere )
	End If

	'Luciano T. Alberti - SMS 90450 - 09/04/2008 - Início
	If Len(vsCampoTipoData) > 0 Then
		sql.ParamByName("CPOTPDATA").AsDateTime =  DateValue(vsCampoTipoData)
	End If
	'Luciano T. Alberti - SMS 90450 - 09/04/2008 - Início

	sql.Active = True

	piRegistros = sql.FieldByName("REG").AsInteger

	sql.Active = False

    Dim xml As String

    If piRegistros <= 50 Then
		sql.Clear
		sql.Add( "SELECT " + psCampos)
		sql.Add( "    FROM " + psFrom )

		If Not ( psWhere = "" ) Then
			psWhere = Replace( Replace( psWhere, "&lt", ">" ), "&gt", "<")
			sql.Add( "WHERE " + psWhere )
		End If

		If Not ( psOrderby = "" ) Then sql.Add( "ORDER BY " + psOrderby )

		'Luciano T. Alberti - SMS 90450 - 09/04/2008 - Início
		If Len(vsCampoTipoData) > 0 Then
			sql.ParamByName("CPOTPDATA").AsDateTime =  DateValue(vsCampoTipoData)
		End If
		'Luciano T. Alberti - SMS 90450 - 09/04/2008 - Início

		sql.Active = True


		xml = ""

		xml = xml + "<resultados>"

		While Not sql.EOF
			xml = xml + "<resultado>"

			For i = 0 To sql.FieldCount - 1
				xml = xml + addXML( sql.Fields( i ).Name, sql.Fields( i ).Name, sql )
			Next i

			xml = xml + "</resultado>"

			sql.Next
		Wend

		xml = xml + "</resultados>"

		Set sql = Nothing
    End If

	ServiceVar("piRegistros") =   CStr( piRegistros )
	ServiceVar("psMensagem") = psMensagem
	ServiceResult = CStr( xml )

	Exit Sub

	erro:
	piRegistros = -1
	psMensagem = sql.Text + "ERROR: " + Err.Description

	ServiceVar("piRegistros") = CStr( piRegistros )
	ServiceVar("psMensagem") = CStr( psMensagem )
End Sub
