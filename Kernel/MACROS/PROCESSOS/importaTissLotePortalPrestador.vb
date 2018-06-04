'HASH: AECA648C8DDFF138EB92E37AAB1400D9
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub Main

  Dim sql As Object

  If (InStr(SQLServer, "MSSQL") > 0) Then
    CriaTabelaTemporariaSqlServer
  End If

  Set sql = NewQuery

  sql.Clear
  sql.Add("SELECT M.HANDLE                     					")
  sql.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS M					")
  sql.Add(" WHERE HANDLE in ("+SessionVar("HANDLETISS")+") 		")
  sql.Active = True

  Dim dllValidarMensagem As Object
  Set dllValidarMensagem = CreateBennerObject("Benner.Saude.WSTiss.Versionador.VersionadorImportarMensagemTISS")

  While Not sql.EOF

	SessionVar("HANDLE") = sql.FieldByName("HANDLE").AsInteger
    SessionVar("HANDLETABELA_TISS") = sql.FieldByName("HANDLE").AsInteger
    SessionVar("NOMETABELA_TISS") = "SAM_PRESTADOR_MENSAGEMTISS"
    SessionVar("NOMECAMPO_TISS") = "ARQUIVORECEBIDO"
	SessionVar("HANDLE_TISVERSAO") = "0"

	dllValidarMensagem.Exec(CurrentSystem)
	sql.Next
  Wend

  Set dllValidarMensagem = Nothing
  Set sql = Nothing


End Sub
