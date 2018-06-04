'HASH: 08E63C1F2DC09A94F2C85305442D79E6
'#Uses "*bsShowMessage"
Option Explicit

Public Sub BOTAOFORNECEDORVENCEDOR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT COUNT(1) QTD")
  SQL.Add("  FROM SAM_LICITACAO_FORNECEDOR ")
  SQL.Add(" WHERE LICITACAO = :LICITACAO")
  SQL.Add("   AND FORNECEDORVENCEDOR = 'S'")
  SQL.Add("   AND HANDLE <> :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("LICITACAO").AsInteger = CurrentQuery.FieldByName("LICITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("QTD").AsInteger > 0 Then
    BsShowMessage("Para esta cotação existe um fornecedor vencedor. É necessário primeiro desqualificá-lo e posteriormente voltar a marcar este fornecedor como vencedor.", "I")
  Else
    SQL.Clear
    SQL.Add("UPDATE SAM_LICITACAO_FORNECEDOR SET FORNECEDORVENCEDOR = 'S',")
    SQL.Add("USUARIOFORNVENCEDOR = :USUARIO, DATAFORNVENCEDOR = :DATA, USUARIODESQUALIFICACAO = NULL, MOTIVODESQUALIFICACAO = NULL, DATADESQUALIFICACAO = NULL, OBSERVACAO = NULL")
    SQL.Add("WHERE HANDLE = :HANDLE")
    SQL.ParamByName("USUARIO").AsInteger = CurrentUser
    SQL.ParamByName("DATA").AsDateTime = ServerNow
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
  End If

  RefreshNodesWithTable("SAM_LICITACAO_FORNECEDOR")

  Set SQL = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT SUM(A.PRECOUNITARIO*D.QUANTIDADE) TOTAL")
  SQL.Add("  FROM SAM_LICITACAO_FORNECEDOR_ITEM A")
  SQL.Add("  JOIN SAM_LICITACAO_FORNECEDOR      B ON (A.LICITACAOFORNECEDOR = B.HANDLE)")
  SQL.Add("  JOIN SAM_LICITACAO                 C ON (B.LICITACAO = C.HANDLE)")
  SQL.Add("  JOIN SAM_LICITACAO_ITENS           D ON (D.LICITACAO = C.HANDLE AND A.ITEM = D.HANDLE)")
  SQL.Add(" WHERE B.HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  LBLVALORTOTAL.Text= "Valor Total Cotado: "+ Format(SQL.FieldByName("TOTAL").AsString, "###,###,##0.00")

  Set SQL = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("MOTIVODESQUALIFICACAO").IsNull Then
    CurrentQuery.FieldByName("DATADESQUALIFICACAO").AsDateTime = ServerNow
    CurrentQuery.FieldByName("USUARIODESQUALIFICACAO").AsInteger = CurrentUser
    CurrentQuery.FieldByName("FORNECEDORVENCEDOR").AsString = "N"
  Else
    CurrentQuery.FieldByName("DATADESQUALIFICACAO").Clear
    CurrentQuery.FieldByName("USUARIODESQUALIFICACAO").Clear
  End If

  If CurrentQuery.State = 3 Then
    'Somente deve gravar quando for inclusão
    CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = CurrentUser
    CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime = ServerNow

    Dim vAno As Integer
    Dim vContador As Long

    vAno = Year(CurrentQuery.FieldByName("DATAHORAENTREGA").AsDateTime)
    NewCounter("WEB_LICITACAO",vAno,1,vContador)
    CurrentQuery.FieldByName("PROTOCOLO").AsString = Trim(Str(vAno))+Trim(Format(vContador,"0000")) 'Trim(Str(vAno)+Str(vContador))
  End If


End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "PARTICIPARLICITACAO" Then
    Dim vContador As Long
    Dim vAno As Integer
    CurrentQuery.Edit
    CurrentQuery.FieldByName("DATAHORAENTREGA").AsDateTime = ServerNow
    vAno = Year(CurrentQuery.FieldByName("DATAHORAENTREGA").AsDateTime)
    NewCounter("WEB_LICITACAO",vAno,1,vContador)
    CurrentQuery.FieldByName("PROTOCOLO").AsString = Str(vAno)+Str(vContador)
    CurrentQuery.Post
    InfoDescription = "Por favor, anote o número do protocolo: " + CurrentQuery.FieldByName("PROTOCOLO").AsString
  ElseIf CommandID = "BOTAOFORNECEDORVENCEDOR" Then
    BOTAOFORNECEDORVENCEDOR_OnClick
  End If

End Sub
