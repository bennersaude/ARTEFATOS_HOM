'HASH: C4BE0FA49D5BD97560397099BFF5E071
Public Sub PUBLICAR_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishGroup(CurrentSystem, -1, CurrentQuery.FieldByName("CODIGO").AsString) 
  Set Obj = Nothing 
End Sub 
 
Public Sub DUPLICAR_OnClick() 
  Dim NewGroupHandle As Integer 
  Dim GroupHandle As Integer 
  Dim GroupName As String 
  Dim GroupCodigo As String 
  Dim ItemHandle As Integer 
  Dim ItemCodigo As String 
  Dim NewItemHandle As Integer 
  Dim FieldsToChange As String 
 
  StartTransaction 
 
  GroupHandle = CurrentQuery.FieldByName("HANDLE").AsInteger 
  GroupName = CurrentQuery.FieldByName("NOME").AsString & "_COPIA" 
  GroupCodigo = CurrentQuery.FieldByName("CODIGO").AsString & "_" & CStr(CurrentSystem.NewHandle("W_GRUPOS")) 
  FieldsToChange = "NOME=" & GroupName & ",CODIGO=" & GroupCodigo 
  If (Not IsServerDeveloper) Then 
    FieldsToChange = FieldsToChange & ",SISTEMA=N" 
  End If 
 
  NewGroupHandle = CopyRecord("W_GRUPOS", GroupHandle, FieldsToChange) 
 
  Dim QGroupItens As Object 
  Set QGroupItens = NewQuery 
  QGroupItens.Text = "SELECT HANDLE, CODIGO FROM W_GRUPOITENS " 
  QGroupItens.Text = QGroupItens.Text & "WHERE GRUPO=" & CStr(GroupHandle) 
  QGroupItens.Text = QGroupItens.Text & " AND ITEM IS NULL" 
  QGroupItens.Active = True 
  Dim QGroupItemGroups As Object 
  Set QGroupItemGroups = NewQuery 
  Dim Fields As String 
 
  While Not QGroupItens.EOF 
    ItemHandle =  QGroupItens.FieldByName("HANDLE").AsInteger 
    ItemCodigo = QGroupItens.FieldByName("CODIGO").AsString 
    NewItemHandle = CopyRecord("W_GRUPOITENS", ItemHandle, "GRUPO="& CStr(NewGroupHandle)& ",CODIGO=" & ItemCodigo) 
 
    'Copia permissoes 
    QGroupItemGroups.Active = False 
    QGroupItemGroups.Text = "SELECT HANDLE FROM W_GRUPOITEMGRUPOS " 
    QGroupItemGroups.Text = QGroupItemGroups.Text & "WHERE ITEM=" & CStr(ItemHandle) 
    QGroupItemGroups.Active = True 
    While Not QGroupItemGroups.EOF 
      Fields = "ITEM="& CStr(NewItemHandle) 
      CopyRecord("W_GRUPOITEMGRUPOS", QGroupItemGroups.FieldByName("HANDLE").AsInteger , Fields) 
      QGroupItemGroups.Next 
    Wend 
 
    CopyItensGroup GroupHandle, NewGroupHandle, ItemHandle, NewItemHandle 
    QGroupItens.Next 
  Wend 
 
  Commit 
 
  RefreshNodesWithTable("W_GRUPOS") 
End Sub 
 
Public Sub CopyItensGroup(aGroupHandle As Integer, aNewGroupHandle As Integer, aItemHandle As Integer, aNewItemHandle As Integer) 
  Dim CurrentNewItemHandle As Integer 
  Dim CurrentItemHandle As Integer 
  Dim Fields As String 
  Dim QGroupItens As Object 
  Set QGroupItens = NewQuery 
  QGroupItens.Text = "SELECT HANDLE FROM W_GRUPOITENS " 
  QGroupItens.Text = QGroupItens.Text & "WHERE GRUPO=" & CStr(aGroupHandle) 
  QGroupItens.Text = QGroupItens.Text & " AND ITEM=" & CStr(aItemHandle) 
  QGroupItens.Active = True 
 
  Dim QGroupItemGroups As Object 
  Set QGroupItemGroups = NewQuery 
  Dim FieldsGroup As String 
 
  While Not QGroupItens.EOF 
    CurrentItemHandle =  QGroupItens.FieldByName("HANDLE").AsInteger 
    Fields = "GRUPO="& CStr(aNewGroupHandle) & ",ITEM="& CStr(aNewItemHandle) 
    CurrentNewItemHandle = CopyRecord("W_GRUPOITENS", CurrentItemHandle, Fields) 
 
    'Copia permissoes 
    QGroupItemGroups.Active = False 
    QGroupItemGroups.Text = "SELECT HANDLE FROM W_GRUPOITEMGRUPOS " 
    QGroupItemGroups.Text = QGroupItemGroups.Text & "WHERE ITEM=" & CStr(CurrentItemHandle) 
    QGroupItemGroups.Active = True 
    While Not QGroupItemGroups.EOF 
      FieldsGroup = "ITEM="& CStr(CurrentNewItemHandle) 
      CopyRecord("W_GRUPOITEMGRUPOS", QGroupItemGroups.FieldByName("HANDLE").AsInteger , FieldsGroup) 
      QGroupItemGroups.Next 
    Wend 
 
    CopyItensGroup aGroupHandle, aNewGroupHandle, CurrentItemHandle, CurrentNewItemHandle 
    QGroupItens.Next 
  Wend 
End Sub 
 
Public Function LimpaStr(Value As String) 
  Caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_" 
  LimpaStr = "" 
 
  For i = 1 To Len(Value) 
    If InStr(Caracteres, Mid(Value, i, 1)) Then 
      LimpaStr = LimpaStr + Mid(Value, i, 1) 
    End If 
  Next i 
End Function 
 
Public Sub CODIGO_OnExit() 
Dim Obj As Object 
  If (CurrentQuery.State <> 1) And (CStr(CurrentQuery.FieldByName("CODIGO").OldValue) <> CStr(CurrentQuery.FieldByName("CODIGO").NewValue)) Then 
    On Error GoTo Fim 
    Set Obj = CreateBennerObject("Pyxis.Helper") 
    CurrentQuery.FieldByName("CODIGO").AsString = Obj.AdjustCodeLevel(CurrentSystem, CurrentQuery.FieldByName("CODIGO").AsString) 
    Set Obj = Nothing 
    Fim: 
    CurrentQuery.FieldByName("CODIGO").AsString = LimpaStr(CurrentQuery.FieldByName("CODIGO").AsString) 
  End If 
End Sub 
 
Public Sub TABLE_AfterInsert() 
  If (IsServerDeveloper) Then 
    CurrentQuery.FieldByName("SISTEMA").AsBoolean = True 
  End If 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  If (Not IsServerDeveloper) Then 
    EhDeSistema = CurrentQuery.FieldByName("SISTEMA").AsBoolean 
    CODIGO.ReadOnly = EhDeSistema 
    NOME.ReadOnly = EhDeSistema 
    AUTOSELECAO.ReadOnly = EhDeSistema 
    OBSERVACOES.ReadOnly = EhDeSistema 
    SISTEMA.ReadOnly = EhDeSistema 
    SISTEMA.Visible = Not EhDeSistema 
  End If 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  If CurrentQuery.State = 3 Then 
    CurrentQuery.FieldByName("ULTIMAALTERACAO").AsDateTime = ServerNow 
  End If 
End Sub 
 
Public Sub TABLE_OnDeleteBtnClick(CanContinue As Boolean) 
  If (Not IsServerDeveloper) And (CurrentQuery.FieldByName("SISTEMA").AsBoolean) Then 
    CanContinue = False 
    MsgBox("Menu não pode ser excluído porque é um menu de sistema") 
  End If 
End Sub 
 
Public Sub TABLE_UpdateRequired() 
  CODIGO_OnExit 
End Sub 
