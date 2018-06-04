'HASH: 755AA28D260B7DC0D2E7E2BAC935A28E
Option Explicit

Public Sub TABLE_AfterInsert()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT TABTIPOAUTEXTERNO FROM SAM_PARAMETROSATENDIMENTO")
  SQL.Active = True
  If SQL.FieldByName("TABTIPOAUTEXTERNO").AsInteger >0 Then
    CurrentQuery.FieldByName("TABTIPOREGISTRO").Value = SQL.FieldByName("TABTIPOAUTEXTERNO").AsInteger -1
  End If
  CurrentQuery.FieldByName("SITUACAO").Value = "A"
  CurrentQuery.FieldByName("DATAINICIAL").Value = ServerNow
  Set SQL = Nothing
End Sub

Public Sub TABLE_AfterPost()

  'No afterpost ainda está em transação.O commit vem depois.
  'Sendo assim,é necessário terminar a transação,antes do processo e abrir a transação depois.

  If InTransaction Then Commit

  If CurrentQuery.FieldByName("TABTIPOREGISTRO").AsInteger = 1 Then

    Dim interface As Object
    Set interface = CreateBennerObject("BSATE001.Importar")

    interface.inicializar(CurrentSystem)

    If(CurrentQuery.FieldByName("TABTIPO").AsInteger = 1)Or(CurrentQuery.FieldByName("TABTIPO").AsInteger = 2)Then
    interface.exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ARQUIVO").AsString)
  End If

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 3 Then
    interface.FechamentoGuias(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("GUIAINICIAL").AsFloat, CurrentQuery.FieldByName("GUIAFINAL").AsFloat, CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime, 0, 0)'CurrentQuery.FieldByName("PEG").AsInteger)
  End If

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 4 Then
    interface.CancelamentoAutorizacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("CHAVE").AsFloat)
  End If

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 5 Then
    interface.Exclusao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("AUTORIZACAOGUIA").AsString, CurrentQuery.FieldByName("CHAVE").AsFloat)
  End If
  interface.finalizar

  Set interface = Nothing

End If

If CurrentQuery.FieldByName("TABTIPOREGISTRO").AsInteger = 2 Then
  Dim interface2 As Object
  Set interface2 = CreateBennerObject("BSATE002.Rotinas")
  interface2.ImportarArquivo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ARQUIVO").AsString)
  Set interface2 = Nothing
End If


If Not InTransaction Then StartTransaction

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object
  Set SQL = NewQuery

  If CurrentQuery.FieldByName("TABTIPOREGISTRO").AsInteger = 1 Then

    If(CurrentQuery.FieldByName("TABTIPO").AsInteger <>1)And _
       (CurrentQuery.FieldByName("TABTIPO").AsInteger <>2)And _
       (CurrentQuery.FieldByName("TABTIPO").AsInteger <>3)And _
       (CurrentQuery.FieldByName("TABTIPO").AsInteger <>4)And _
       (CurrentQuery.FieldByName("TABTIPO").AsInteger <>5)Then
    CanContinue = False
    CancelDescription = "Erro: TABTIPO inválido"
  End If

  If(CurrentQuery.FieldByName("TABTIPO").AsInteger = 1)Or(CurrentQuery.FieldByName("TABTIPO").AsInteger = 2)Then
  If Len(CurrentQuery.FieldByName("ARQUIVO").AsString) = 0 Then
    CanContinue = False
    CancelDescription = "Erro: ARQUIVO não informado"
  End If
End If

If CurrentQuery.FieldByName("TABTIPO").AsInteger = 3 Then
  If(CurrentQuery.FieldByName("GUIAINICIAL").IsNull)Or(CurrentQuery.FieldByName("GUIAINICIAL").IsNull)Then
  CanContinue = False
  CancelDescription = "Guia inicial ou final não informada"
End If
If(CurrentQuery.FieldByName("DATARECEBIMENTO").IsNull)Then
CanContinue = False
CancelDescription = "Data de recebimento não informada"
End If
End If

If CurrentQuery.FieldByName("TABTIPO").AsInteger = 4 Then
  If(CurrentQuery.FieldByName("CHAVE").IsNull)Then
  CanContinue = False
  CancelDescription = "Autorização/chave não informada"
End If
End If

If CurrentQuery.FieldByName("TABTIPO").AsInteger = 5 Then
  If(CurrentQuery.FieldByName("CHAVE").IsNull)Then
  CanContinue = False
  CancelDescription = "Chave não informada"
End If
If(CurrentQuery.FieldByName("AUTORIZACAOGUIA").IsNull)Then
CanContinue = False
CancelDescription = "Autorização ou guia não informado"
End If
End If

End If



If CurrentQuery.FieldByName("TABTIPOREGISTRO").AsInteger = 2 Then
  If Len(CurrentQuery.FieldByName("ARQUIVO").AsString) = 0 Then
    CanContinue = False
    CancelDescription = "Erro: ARQUIVO não informado"
  End If
End If

Set SQL = Nothing

End Sub

