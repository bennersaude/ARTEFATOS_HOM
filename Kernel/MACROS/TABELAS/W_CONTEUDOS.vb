'HASH: E34B972BFBE538D782FA6D15FF08F7CF
Dim TituloOriginal As String 
Dim GrupoOriginal As Integer 
 
Public Sub TABLE_AfterEdit() 
  TituloOriginal = CurrentQuery.FieldByName("TITULO").AsString 
  GrupoOriginal = CurrentQuery.FieldByName("GRUPO").AsInteger 
End Sub 
 
Public Sub TABLE_AfterInsert() 
Dim iGrupo As Long 
  TituloOriginal = "" 
  GrupoOriginal = 0 
  iGrupo = RecordHandleOfTable("W_CONTEUDOS|NS") 
  If iGrupo > 0 Then 
    CurrentQuery.FieldByName("GRUPO").AsInteger = iGrupo 
  End If 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  CODIGOINTERNO.Text = "Código interno: " + CurrentQuery.FieldByName("HANDLE").AsString 
End Sub 
 
Public Sub TABLE_UpdateRequired() 
  'Para evitar problemas na gravação, inicializa conteúdo do campo IDENTIFICACAO 
  If (CurrentQuery.FieldByName("IDENTIFICACAO").IsNull) Then 
    CurrentQuery.FieldByName("IDENTIFICACAO").AsString = CurrentQuery.FieldByName("TITULO").AsString 
  End If 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
Dim qWork As Object 
 
  If (TituloOriginal <> CurrentQuery.FieldByName("TITULO").AsString) Or (GrupoOriginal <> CurrentQuery.FieldByName("GRUPO").AsInteger) Then 
 
    'Verifica se está fazendo uma referência circular e pega os títulos dos níveis ancestrais 
    Set qWork = NewQuery 
    s = CurrentQuery.FieldByName("TITULO").AsString 
    iGrupo = CurrentQuery.FieldByName("GRUPO").AsInteger 
    While iGrupo <> 0 
 
      If (iGrupo = CurrentQuery.FieldByName("HANDLE").AsInteger) Then 
        If (VisibleMode) Then 
          MsgBox("Grupo informado é subitem deste conteúdo, gerando uma referência circular.") 
        Else 
          CancelDescription = "Grupo informado é subitem deste conteúdo, gerando uma referência circular." 
        End If 
        CanContinue = False 
        Exit Sub 
      End If 
 
      qWork.Clear 
      qWork.Add("SELECT TITULO, GRUPO FROM W_CONTEUDOS WHERE HANDLE = " + CStr(iGrupo)) 
      qWork.Active = True 
      iGrupo = qWork.FieldByName("GRUPO").AsInteger 
      s = qWork.FieldByName("TITULO").AsString + " > " + s 
      qWork.Active = False 
    Wend 
    CurrentQuery.FieldByName("IDENTIFICACAO").AsString = s 
  End If 
  Set qWork = Nothing 
End Sub 
 
Public Sub TABLE_AfterPost() 
  If (TituloOriginal <> CurrentQuery.FieldByName("TITULO").AsString) Or (GrupoOriginal <> CurrentQuery.FieldByName("GRUPO").AsInteger) Then 
    'Atualiza o identificador dos descendentes, caso seja necessário 
    AtualizaIdentificacao CurrentQuery.FieldByName("IDENTIFICACAO").AsString, CurrentQuery.FieldByName("HANDLE").AsInteger 
  End If 
 
  CODIGOINTERNO.Text = "Código interno: " + CurrentQuery.FieldByName("HANDLE").AsString 
End Sub 
 
Private Sub AtualizaIdentificacao(ByVal aIdentificacao As String, ByVal aGrupo As Integer) 
Dim qEdit As BPesquisa 
  Set qEdit = NewQuery 
  qEdit.RequestLive = True 
  qEdit.Text = "SELECT HANDLE, TITULO, IDENTIFICACAO FROM W_CONTEUDOS WHERE GRUPO = " + CStr(aGrupo) 
  qEdit.Active = True 
 
  While Not qEdit.EOF 
    qEdit.Edit 
    s = aIdentificacao + " > " + qEdit.FieldByName("TITULO").AsString 
    qEdit.FieldByName("IDENTIFICACAO").AsString = s 
    qEdit.Post 
    AtualizaIdentificacao s, qEdit.FieldByName("HANDLE").AsInteger 
    qEdit.Next 
  Wend 
 
  qEdit.Active = False 
  Set qEdit = Nothing 
End Sub 
