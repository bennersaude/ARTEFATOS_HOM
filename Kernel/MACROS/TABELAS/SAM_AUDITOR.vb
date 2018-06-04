'HASH: A703D9AABAA8E802CB939F2917404AD3

'SAM_AUDITOR
'JULIANA 12/04/2002

'# uses "*ProcuraPrestador"
'#Uses "*bsShowMessage"

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("N", "T", PRESTADOR.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If


End Sub

Public Sub USUARIO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CODIGO|APELIDO|NOME"

  vCriterio = ""

  vCampos = "Código|Apelido|Nome"

  vHandle = interface.Exec(CurrentSystem, "Z_GRUPOUSUARIOS", vColunas, 3, vCampos, vCriterio, "Usuários", True, USUARIO.Text)

  CurrentQuery.Edit
  CurrentQuery.FieldByName("USUARIO").Value = vHandle

  ShowPopup = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)  ' SMS 91312 - Paulo Melo - 12/01/2008 - Não deixar cadastrar dois registros com
																							 ' a mesma identificacao
	Dim q As Object
	Set q = NewQuery

' PROTEÇÃO PARA NÃO TER 2 REGISTROS COM MESMA IDENTIFICAÇÃO
	q.Add("SELECT HANDLE")
	q.Add("FROM SAM_AUDITOR")
	q.Add("WHERE IDENTIFICACAO = :IDENTIFICACAO")

	If CurrentQuery.State = 2 Then  'modo de edição, pois na web o handle não deve ser levado em conta na hora da inserção
		q.Add("AND HANDLE <> :HANDLE")
	End If

	q.ParamByName("IDENTIFICACAO").AsInteger = CurrentQuery.FieldByName("IDENTIFICACAO").AsInteger

	If CurrentQuery.State = 2 Then
		q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	End If

	q.Active = True

	If q.FieldByName("HANDLE").AsInteger > 0 Then
		bsShowMessage("Identificação já cadastrada!", "E")
	 	CanContinue = False
		Set q = Nothing
		Exit Sub
	End If

	Set q = Nothing

	'----------------------------- PROTEÇÃO PARA NÃO TER 2 REGISTROS COM MESMO NOME		' SMS 94734 - Paulo Melo - 19/03/2008 - Não cadastrar nome de perito duplicado
	Dim q1 As Object
	Set q1 = NewQuery

	q1.Add("SELECT NOME")
	q1.Add("FROM SAM_AUDITOR")
	q1.Add("WHERE HANDLE <> :HANDLE")
	q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	q1.Active = True

	While Not q1.EOF
		If UCase(Trim(q1.FieldByName("NOME").AsString)) = UCase(Trim(CurrentQuery.FieldByName("NOME").AsString)) Then
			bsShowMessage("Nome já cadastrado!", "E")
			CanContinue = False
			Set q = Nothing
			Exit Sub
		End If
		q1.Next
	Wend

	Set q1 = Nothing

End Sub
