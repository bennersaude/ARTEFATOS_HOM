'HASH: 73B6D963AAAC293421F8B8D983788111
'#Uses "*bsShowMessage"

Public Sub BOTAOGERAR_OnClick()
  Dim vRECEITA As String
  vRECEITA = ""

  If Not CurrentQuery.FieldByName("TEXTO").IsNull Then
    Dim RTF2TXTDLL As Object
    Set RTF2TXTDLL = CreateBennerObject("RTF2TXT.ROTINAS")
    Dim TXT As Object
    Set TXT = NewQuery

    TXT.Clear
    TXT.Add("SELECT TEXTO FROM CLI_TEXTO")
    TXT.Add(" WHERE HANDLE = :HANDLE")
    TXT.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TEXTO").AsInteger
    TXT.Active = True

    vRECEITA = RTF2TXTDLL.Rtf2Txt(CurrentSystem, TXT.FieldByName("TEXTO").AsString) + Chr(13) + Chr(13)

    Set TXT = Nothing
    Set RTF2TXTDLL = Nothing
  End If

  Dim medicamento As Object
  Set medicamento = NewQuery

  medicamento.Clear
  medicamento.Add("SELECT M.DESCRICAO, A.DESCRICAO APRESENTACAO, RM.QUANTIDADE")
  medicamento.Add("  FROM CLI_RECEITUARIOMATMED RM,")
  medicamento.Add("       SAM_MATMED M,")
  medicamento.Add("       SAM_MATMEDBRAPRESENTACAO A")
  medicamento.Add(" WHERE RECEITUARIO = :RECEITUARIO")
  medicamento.Add("   AND RM.MATMED = M.HANDLE")
  medicamento.Add("   AND M.BRAPRESENTACAO = A.HANDLE")
  medicamento.ParamByName("RECEITUARIO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  medicamento.Active = True

  While Not medicamento.EOF
    vRECEITA = vRECEITA + medicamento.FieldByName("QUANTIDADE").AsString + "  " + medicamento.FieldByName("DESCRICAO").AsString + Chr(13)
    vRECEITA = vRECEITA + medicamento.FieldByName("APRESENTACAO").AsString + Chr(13) + Chr(13)
    medicamento.Next
  Wend

  Set medicamento = Nothing

  CurrentQuery.Edit
  CurrentQuery.FieldByName("RECEITA").AsString = vRECEITA
  CurrentQuery.Post

  RefreshNodesWithTable("CLI_RECEITUARIO")
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()

  Dim RelatorioHandle As Long
  Dim QueryBuscaHandle As Object
  Set QueryBuscaHandle = NewQuery

  QueryBuscaHandle.Active = False
  QueryBuscaHandle.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'CLI012'")
  QueryBuscaHandle.Active = True

  RelatorioHandle = QueryBuscaHandle.FieldByName("HANDLE").AsInteger

  Set QueryBuscaHandle = Nothing

  On Error GoTo Final

  UserParam = "" 'se userparam for vazio ele imprime a carga corrente senão ele imprime o receituario do handle
  'passado com parâmetro pelo sistema na tela de agendamento

  ReportPreview(RelatorioHandle, "", True, False)
  Exit Sub

Final :
  bsShowMessage(Str(Error), "E")
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT A.RECURSO FROM CLI_ATENDIMENTO A WHERE A.HANDLE = :ATENDIMENTO")
  SQL.Add("AND EXISTS(SELECT 1 FROM CLI_RECURSO_USUARIO RU, CLI_RECURSO R WHERE R.HANDLE = A.RECURSO AND R.PRESTADOR = RU.PRESTADOR AND RU.USUARIO = :USUARIO)")
  SQL.ParamByName("ATENDIMENTO").AsInteger = RecordHandleOfTable("CLI_ATENDIMENTO")
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True
  If SQL.EOF Then
    CanContinue = False
    bsShowMessage("Usuário inválido!", "E")
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOGERAR") Then
		BOTAOGERAR_OnClick
	End If
	If (CommandID = "BOTAOIMPRIMIR") Then
		BOTAOIMPRIMIR_OnClick
	End If
End Sub
