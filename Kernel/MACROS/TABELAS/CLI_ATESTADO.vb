'HASH: A2DD8A8F4A0C911B8A1B57372687BA5A
'#Uses "*bsShowMessage"

Public Sub BOTAOIMPRIMIR_OnClick()

  Dim RelatorioHandle As Long
  Dim QueryBuscaHandle As Object
  Set QueryBuscaHandle = NewQuery

  QueryBuscaHandle.Active = False
  QueryBuscaHandle.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'CLI002'")
  QueryBuscaHandle.Active = True

  RelatorioHandle = QueryBuscaHandle.FieldByName("HANDLE").AsInteger

  Set QueryBuscaHandle = Nothing

  On Error GoTo Final

  UserParam = "" 'no relatório no momento de selecionar o registro para impressão ele verifica se foi passado
  'algum parâmetro,caso afirmativo significa que foi pedido uma impressão pela tela de atendimento
  'e o valor do handle do registro está presente em UserParam
  ReportPreview(RelatorioHandle, "", True, False)
  Exit Sub

Final :
  bsShowMessage(Str(Error), "E")
End Sub

Public Sub BOTAOMONTAR_OnClick()
  Dim vRECEITA As String
  vRECEITA = ""

  If Not CurrentQuery.FieldByName("TEXTOPADRAO").IsNull Then
    Dim RTF2TXTDLL As Object
    Set RTF2TXTDLL = CreateBennerObject("RTF2TXT.ROTINAS")
    Dim TXT As Object
    Set TXT = NewQuery

    TXT.Clear
    TXT.Add("SELECT TEXTO FROM CLI_TEXTO")
    TXT.Add(" WHERE HANDLE = :HANDLE")
    TXT.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TEXTOPADRAO").AsInteger
    TXT.Active = True

    vRECEITA = RTF2TXTDLL.Rtf2Txt(CurrentSystem, TXT.FieldByName("TEXTO").AsString) + Chr(13) + Chr(13)

    Set TXT = Nothing
    Set RTF2TXTDLL = Nothing
  End If

  If CurrentQuery.State = 1 Then
    CurrentQuery.Edit
  End If

  CurrentQuery.FieldByName("TEXTO").AsString = vRECEITA
  'CurrentQuery.Post

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
	If (CommandID = "BOTAOIMPRIMIR") Then
		BOTAOIMPRIMIR_OnClick
	End If
	If (CommandID = "BOTAOMONTAR") Then
		BOTAOMONTAR_OnClick
	End If
End Sub
