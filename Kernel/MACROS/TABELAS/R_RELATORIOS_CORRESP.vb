'HASH: A58E6B3E20D2086D149B1B1DD7A10DDB
'#Uses "*bsShowMessage"

Option Explicit

Dim SQLProcuraRelatorio As Object

Public Sub RELATORIO_OnBtnClick()
  Dim Interface As Object
  Dim HandleFiltro As Long
  Dim SQLProcuraRelatorio As Object
  Set Interface = CreateBennerObject("Procura.Procurar")
  HandleFiltro = Interface.Exec(CurrentSystem, "R_RELATORIOS", "CODIGO|NOME", 2, "Código|Relatório", "CONTROLACORRESP = 'S'", "Procura Relatórios", True, "")
  Set Interface = Nothing
  If HandleFiltro >0 Then
    'Busca pelo RFFiltro selecionado
    Set SQLProcuraRelatorio = NewQuery
    SQLProcuraRelatorio.Add("SELECT CODIGO")
    SQLProcuraRelatorio.Add("FROM R_RELATORIOS")
    SQLProcuraRelatorio.Add("WHERE HANDLE =" + Str(HandleFiltro))

    SQLProcuraRelatorio.Active = False
    SQLProcuraRelatorio.Active = True

    CurrentQuery.Edit
    CurrentQuery.FieldByName("RELATORIO").AsString = SQLProcuraRelatorio.FieldByName("CODIGO").AsString

    'Mata a query
    Set SQLProcuraRelatorio = Nothing

  End If

End Sub

Public Sub TABLE_AfterPost()
  Set SQLProcuraRelatorio = NewQuery

  SQLProcuraRelatorio.Add("SELECT NOME")
  SQLProcuraRelatorio.Add("FROM R_RELATORIOS")
  SQLProcuraRelatorio.Add("WHERE CODIGO =:P_CODIGO")

  SQLProcuraRelatorio.ParamByName("P_CODIGO").Value = CurrentQuery.FieldByName("RELATORIO").AsString
  SQLProcuraRelatorio.Active = False
  SQLProcuraRelatorio.Active = True

  NOMERELATORIO.Text = SQLProcuraRelatorio.FieldByName("NOME").AsString

  Set SQLProcuraRelatorio = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Set SQLProcuraRelatorio = NewQuery

  SQLProcuraRelatorio.Add("SELECT NOME")
  SQLProcuraRelatorio.Add("FROM R_RELATORIOS")
  SQLProcuraRelatorio.Add("WHERE CODIGO =:P_CODIGO")

  SQLProcuraRelatorio.ParamByName("P_CODIGO").Value = CurrentQuery.FieldByName("RELATORIO").AsString
  SQLProcuraRelatorio.Active = False
  SQLProcuraRelatorio.Active = True

  NOMERELATORIO.Text = SQLProcuraRelatorio.FieldByName("NOME").AsString

  Set SQLProcuraRelatorio = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Set SQLProcuraRelatorio = NewQuery

  If CurrentQuery.State <>2 Then
    'Verifica se o Relatório já está cadastrado
    SQLProcuraRelatorio.Add("SELECT RELATORIO")
    SQLProcuraRelatorio.Add("FROM R_RELATORIOS_CORRESP")
    SQLProcuraRelatorio.Add("WHERE RELATORIO =:P_CODIGO")

    SQLProcuraRelatorio.ParamByName("P_CODIGO").Value = CurrentQuery.FieldByName("RELATORIO").AsString
    SQLProcuraRelatorio.Active = False
    SQLProcuraRelatorio.Active = True

    If Not SQLProcuraRelatorio.EOF Then
      bsShowMessage("Este relatório já está cadastrado controle de Correspondência!", "E")

      CanContinue = False

    End If

  End If

  'Verifica se o relatório possui controle de correspondência
  SQLProcuraRelatorio.Clear

  SQLProcuraRelatorio.Add("SELECT CONTROLACORRESP")
  SQLProcuraRelatorio.Add("FROM R_RELATORIOS")
  SQLProcuraRelatorio.Add("WHERE CODIGO =:P_CODIGO")

  SQLProcuraRelatorio.ParamByName("P_CODIGO").Value = CurrentQuery.FieldByName("RELATORIO").AsString
  SQLProcuraRelatorio.Active = False
  SQLProcuraRelatorio.Active = True

  If SQLProcuraRelatorio.FieldByName("CONTROLACORRESP").AsString <>"S" Then
    bsShowMessage("Este relatório não possui controle de Correspondência!", "E")

    CanContinue = False

  End If

  Set SQLProcuraRelatorio = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "RELATORIO") Then
		RELATORIO_OnBtnClick
	End If
End Sub
