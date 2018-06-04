'HASH: 3EA63D6608B31D742D299361B9ABB359
 
'#Uses "*bsShowMessage"
Public Sub BOTAOGERARARQUIVOS_OnClick()

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro não pode estar em edição!", "E")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
    bsShowMessage("Rotina já processada!", "E")
    Exit Sub
  End If

  Dim interface As Object
  Dim vsRetorono As String

  Set interface = CreateBennerObject("Benner.Saude.WSTiss.ImportacaoXmlLote.ImportacaoXmlLote")
  vsRetorno = interface.ImportarXML(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  If vsRetorno = "OK" Then
    BsShowMessage("Processo concluído. Verifique as ocorrências." ,"I")
  Else
    BSShowMessage(vsRetorno, "E")
  End If

  Set interface = Nothing

  If VisibleMode Then
    RefreshNodesWithTable("TIS_IMPORTACAOXMLLOTE")
  End If

End Sub

Public Sub BOTAOPROCESSARARQUIVOS_OnClick()

  Dim viNumArquivos As Integer

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro não pode estar em edição!", "E")
    Exit Sub
  End If

  Dim SQL As Object
  Dim SQLParamGeraisTISS As Object
  Dim SQLArquivos As Object

  Set SQL = NewQuery
  Set SQLParamGeraisTISS = NewQuery
  Set SQLArquivos = NewQuery

  SQLArquivos.Clear
  SQLArquivos.Add("SELECT HANDLE,ARQUIVOXML")
  SQLArquivos.Add("  FROM TIS_IMPORTACAOXMLLOTE_ARQ")
  SQLArquivos.Add(" WHERE IMPORTACAOXMLLOTE = :XMLLOTE")
  SQLArquivos.ParamByName("XMLLOTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLArquivos.Active = True

  If SQLArquivos.EOF Then
    bsShowMessage("Não existem arquivos a serem processados!", "E")
    Exit Sub
  End If

  SQLParamGeraisTISS.Clear
  SQLParamGeraisTISS.Add("SELECT CAMINHOARQUIVOSAGENDADOS FROM TIS_PARAMETROS")
  SQLParamGeraisTISS.Active = True

  If Len(SQLParamGeraisTISS.FieldByName("CAMINHOARQUIVOSAGENDADOS").AsString) = 0 Then
    bsShowMessage("O parâmetro Pasta agendamento TISS deve ser preenchido", "E")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
    bsShowMessage("Rotina já processada!", "E")
    Exit Sub
  End If

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM Z_MACROS WHERE NOME = :NOME")
  SQL.ParamByName("NOME").AsString = "importacaoXMLLote"
  SQL.Active = True

  If Not SQL.EOF Then

    Dim sx As CSServerExec
    Set sx = NewServerExec
    sx.Description = "Processamento em lote de Mensagens TISS"
    sx.Process = SQL.FieldByName("HANDLE").AsInteger
    sx.SessionVar("HANDLEROTINAXMLLOTE") = CurrentQuery.FieldByName("HANDLE").AsString
    sx.Execute

    SQL.Clear
    SQL.Add("UPDATE TIS_IMPORTACAOXMLLOTE SET SITUACAO = :SITUACAO WHERE HANDLE = :HANDLE")
    SQL.ParamByName("SITUACAO").AsInteger = 2
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

    Set sx = Nothing

    bsShowMessage("O processo foi enviado para o servidor. Verifique o monitor de processos.", "I")
  Else
    bsShowMessage("Processo importacaoXMLLote não encontrado na base de dados. Entre em contato com o suporte", "E")
  End If

  Set SQL = Nothing
  Set SQLParamGeraisTISS = Nothing
  Set SQLArquivos = Nothing

  If VisibleMode Then
    RefreshNodesWithTable("TIS_IMPORTACAOXMLLOTE")
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

Dim SQL As Object
Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT COUNT(1) QTD")
SQL.Add("  FROM TIS_IMPORTACAOXMLLOTE_ARQ")
SQL.Add(" WHERE IMPORTACAOXMLLOTE = :XMLLOTE")
SQL.ParamByName("XMLLOTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True

If SQL.FieldByName("QTD").AsInteger > 0 Then
  bsShowMessage("Já existem arquivos carregados para esta rotina. Não é possível excluir!", "E")
  CanContinue = False
End If

Set SQL = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "BOTAOGERARARQUIVOS" Then
	BOTAOGERARARQUIVOS_OnClick
  ElseIf CommandID = "BOTAOPROCESSARARQUIVOS" Then
	BOTAOPROCESSARARQUIVOS_OnClick
  End If
End Sub
