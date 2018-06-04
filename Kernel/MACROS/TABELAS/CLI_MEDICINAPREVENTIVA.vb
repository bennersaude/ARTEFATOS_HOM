'HASH: DFCA56ED0BC09D80BC1AE17974C47D2A
'CLI_MEDICINAPREVENTIVA
'#Uses "*bsShowMessage"

Public Sub BOTAOALTA_OnClick()
  If Not CurrentQuery.FieldByName("USUARIOALTA").IsNull Then
    bsShowMessage("Já foi dada a alta!", "I")
    Exit Sub
  End If

  If CurrentQuery.State = 1 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("DATAALTA").AsDateTime = ServerNow
    CurrentQuery.FieldByName("USUARIOALTA").AsInteger = CurrentUser
    CurrentQuery.Post
    RefreshNodesWithTable("CLI_MEDICINAPREVENTIVA")
  Else
    bsShowMessage("O registro não pode estar em edição ou inserção!", "I")
  End If
End Sub

Public Sub ESPECIALIDADE_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabelas As String

  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_ESPECIALIDADE.CODIGO|SAM_ESPECIALIDADE.DESCRICAO"
  vCriterio = ""
  vCampos = "Código|Especialidade"
  vTabelas = "SAM_ESPECIALIDADE"

  vHandle = Interface.Exec(CurrentSystem, vTabelas, vColunas, 2, vCampos, vCriterio, "Especialidade", False, "")

  If vHandle <>0 Then
    If CurrentQuery.State <>3 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("ESPECIALIDADE").Value = vHandle
  End If
  Set Interface = Nothing
End Sub

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabelas As String

  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_MATRICULA.CPF|SAM_MATRICULA.RG|SAM_MATRICULA.DATANASCIMENTO"
  vCriterio = ""
  vCampos = "Matrícula Funcional|Nome|Codigo|CPF|RG|Dt.Nascimento"
  vTabelas = "SAM_BENEFICIARIO|SAM_MATRICULA[SAM_BENEFICIARIO.MATRICULA = SAM_MATRICULA.HANDLE ]"

  vHandle = Interface.Exec(CurrentSystem, vTabelas, vColunas, 1, vCampos, vCriterio, "Beneficiários", False, "")

  If vHandle <>0 Then
    If CurrentQuery.State <>3 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
  Set Interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim busca As Object
  Set busca = NewQuery
  busca.Add("SELECT HANDLE FROM CLI_MEDICINAPREVENTIVA")
  busca.Add("WHERE BENEFICIARIO = :BENEFICIARIO AND ESPECIALIDADE = :ESPECIALIDADE AND USUARIOALTA IS NULL")
  busca.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  busca.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  busca.Active = True

  If Not busca.EOF Then
    bsShowMessage("O beneficiário já está em tratamento nesta especialidade!", "E")
    CanContinue = False
    Exit Sub
  End If
  Set busca = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOALTA") Then
		BOTAOALTA_OnClick
	End If

End Sub
