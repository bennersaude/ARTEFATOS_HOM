'HASH: DE234BF66A105694D32E72238BF8ED5D
 
'#uses "*CriaTabelaTemporariaSqlServer"

Option Explicit

'--------------------------------------------------------------------------------------------------------------------------
'  SOMENTE USAR A PARTIR DA SAM_AUTORIZ_EVENTOGERADO----------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------

Public Sub reverter(pUsuario As Long)


	CriaTabelaTemporariaSqlServer
	Dim mensagem As String
	Dim dll As Object


	Set dll=CreateBennerObject("samauto.autorizador")
	dll.inicializar(CurrentSystem, "A")
	mensagem = dll.reverterSAM_AUTORIZ_EVENTOGERADO_comMotivo( _
		CurrentSystem, _
		RecordHandleOfTable("SAM_AUTORIZ_EVENTOGERADO"), _
		pUsuario, _
		CurrentQuery.FieldByName("MOTIVOREVERSAO").AsInteger, _
		"G")
	dll.finalizar
	Set dll=Nothing
	If mensagem<>"" Then
		InfoDescription = mensagem
	Else
		InfoDescription = "Reversão concluída com sucesso"
	End If
End Sub


Public Sub TABLE_AfterPost()
	Dim vUsuario As Long
	vUsuario = verificaUsuario
	If vUsuario > 0 Then
		reverter(vUsuario)
	Else
		InfoDescription = "Usuário e senha não conferem"
	End If
End Sub

Public Function verificaUsuario As Long
	' se não digitar o usuário/SENHA, assume o USUARIO corrente
	If (CurrentQuery.FieldByName("USUARIO").AsString <> "") Then
		Dim sql As BPesquisa
		Set sql=NewQuery
		sql.Add("SELECT HANDLE, SENHA FROM Z_GRUPOUSUARIOS WHERE "+SQLUpper+"(APELIDO)=:APELIDO")
		sql.ParamByName("APELIDO").AsString = UCase(CurrentQuery.FieldByName("USUARIO").AsString)
		sql.Active=True
		If sql.FieldByName("SENHA").AsString = PasswordEncode(UCase(CurrentQuery.FieldByName("SENHA").AsString)) Then
			verificaUsuario = sql.FieldByName("HANDLE").AsInteger
		Else
			verificaUsuario = 0
		End If
		Set sql = Nothing
	Else
		verificaUsuario = CurrentUser
	End If
End Function


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim dll As Object
  Set dll=CreateBennerObject("samauto.autorizador")
  If Not dll.verificaNecessidadeReverter(CurrentSystem, RecordHandleOfTable("SAM_AUTORIZ_EVENTOGERADO"), "G") Then
  	CanContinue=False
  	CancelDescription = "O evento não está negado ou está cancelado"
  End If
  Set dll=Nothing
End Sub
