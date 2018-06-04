'HASH: 1DF8D535048263A02148183AF211B198
'CLI_TIPOATENDIMENTO
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|DESCRICAO"
  vCriterio = "ULTIMONIVEL = 'S'"
  vCampos = "Estrutura|Descrição"
  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, vCriterio, "Evento", False, "")

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterPost()
  RefreshNodesWithTable("CLI_TIPOATENDIMENTO")
End Sub

Public Sub TABLE_AfterScroll()
	Dim a As Integer
	Dim Sql As BPesquisa
	Set Sql = NewQuery

	Sql.Clear
	Sql.Add("SELECT TISSTIPOSOLICITACAO FROM SAM_TIPOAUTORIZ WHERE HANDLE = :HANDLE")
	Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
	Sql.Active = True

	TIPOATENDIMENTOTISS.Visible = True
	TIPOATENDIMENTOTISSODONTO.Visible = False

	If (Not Sql.FieldByName("TISSTIPOSOLICITACAO").IsNull) Then
		If (Sql.FieldByName("TISSTIPOSOLICITACAO").AsInteger = 3) Then
			TIPOATENDIMENTOTISS.Visible = False
			TIPOATENDIMENTOTISSODONTO.Visible = True
		End If
	End If
	Set Sql = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim BUSCA As Object
  Set BUSCA = NewQuery
  BUSCA.Add("SELECT * FROM SAM_TIPOAUTORIZ WHERE HANDLE = :TIPOAUTORIZ")
  BUSCA.ParamByName("TIPOAUTORIZ").AsInteger = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
  BUSCA.Active = True

  If BUSCA.FieldByName("FINALIDADEATENDIMENTO").IsNull Then
    bsShowMessage("O tipo de autorização selecionado não possui o campo ""Finalidade de atendimento"" informado!", "E")
    CanContinue = False
    Exit Sub
  ElseIf WebMode Then
    CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger = BUSCA.FieldByName("FINALIDADEATENDIMENTO").AsInteger
  End If

  If BUSCA.FieldByName("LOCALATENDIMENTO").IsNull Then
    bsShowMessage("O tipo de autorização selecionado não possui o campo ""Local de atendimento"" informado!", "E")
    CanContinue = False
    Exit Sub
  ElseIf WebMode Then
    CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger = BUSCA.FieldByName("LOCALATENDIMENTO").AsInteger
  End If

  If BUSCA.FieldByName("REGIMEATENDIMENTO").IsNull Then
    bsShowMessage("O tipo de autorização selecionado não possui o campo ""Regime de atendimento"" informado!", "E")
    CanContinue = False
    Exit Sub
  ElseIf WebMode Then
    CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger = BUSCA.FieldByName("REGIMEATENDIMENTO").AsInteger
  End If

  If BUSCA.FieldByName("CONDICAOATENDIMENTO").IsNull Then
    bsShowMessage("O tipo de autorização selecionado não possui o campo ""Condição de atendimento"" informado!", "E")
    CanContinue = False
    Exit Sub
  ElseIf WebMode Then
    CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger = BUSCA.FieldByName("CONDICAOATENDIMENTO").AsInteger
  End If

  If BUSCA.FieldByName("TIPOTRATAMENTO").IsNull Then
    bsShowMessage("O tipo de autorização selecionado não possui o campo ""Tipo de tratamento"" informado!", "E")
    CanContinue = False
    Exit Sub
  ElseIf WebMode Then
    CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger = BUSCA.FieldByName("TIPOTRATAMENTO").AsInteger
  End If

  If BUSCA.FieldByName("OBJETIVOTRATAMENTO").IsNull Then
    bsShowMessage("O tipo de autorização selecionado não possui o campo ""Objetivo de tratamento"" informado!", "E")
    CanContinue = False
    Exit Sub
  ElseIf WebMode Then
    CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger = BUSCA.FieldByName("OBJETIVOTRATAMENTO").AsInteger
  End If

  Set BUSCA = Nothing
End Sub

Public Sub TIPOAUTORIZACAO_OnChange()
	Dim Sql As BPesquisa
	Set Sql = NewQuery

	Sql.Clear
	Sql.Add("SELECT TISSTIPOSOLICITACAO FROM SAM_TIPOAUTORIZ WHERE DESCRICAO = :DESCRICAO")
	Sql.ParamByName("DESCRICAO").AsString = TIPOAUTORIZACAO.Text
	Sql.Active = True

	If Sql.FieldByName("TISSTIPOSOLICITACAO").AsInteger = 3 Then
		TIPOATENDIMENTOTISS.Visible = False
		TIPOATENDIMENTOTISSODONTO.Visible = True
	Else
		TIPOATENDIMENTOTISS.Visible = True
		TIPOATENDIMENTOTISSODONTO.Visible = False
	End If
	Set Sql = Nothing
End Sub
