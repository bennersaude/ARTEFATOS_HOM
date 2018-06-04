'HASH: B820723110C76762FE5942AD24C0AF87
 
Public Sub EDITARVISAO_OnClick() 
Dim Obj As Object 
  If (CurrentQuery.FieldByName("LINK").AsInteger = 2) Or (CurrentQuery.FieldByName("LINK").AsInteger = 3) Then 
    Set Obj = CreateBennerObject("Pyxis.WebVisionDesigner") 
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("LINKVISAO").AsInteger) 
    Set Obj = Nothing 
  ElseIf (CurrentQuery.FieldByName("LINK").AsInteger = 8) Then 
    Set Obj = CreateBennerObject("Pyxis.WebVisionDesigner") 
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("LINKTABVIRTUAL").AsInteger) 
    Set Obj = Nothing 
  End If 
End Sub 
 
Public Sub LINK_OnChange() 
  CurrentQuery.UpdateRecord 
  EDITARVISAO.Enabled = (CurrentQuery.FieldByName("LINK").AsInteger = 2) Or (CurrentQuery.FieldByName("LINK").AsInteger = 3) Or (CurrentQuery.FieldByName("LINK").AsInteger = 8) 
End Sub 
 
Public Sub PERMISSAOACESSO_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.PermissionConfig") 
  Obj.MenuItem(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub TABLE_AfterPost() 
Dim QWork As BPesquisa 
  Set QWork = NewQuery 
  QWork.Add("UPDATE W_GRUPOS SET ULTIMAALTERACAO = :DATA, USUARIO = :USUARIO WHERE HANDLE = :HANDLE") 
  QWork.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GRUPO").AsInteger 
  QWork.ParamByName("DATA").AsDateTime = ServerNow 
  QWork.ParamByName("USUARIO").AsString = UserNickName 
  QWork.ExecSQL 
  Set QWork = Nothing 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  PERMISSAOACESSO.Enabled = CurrentQuery.FieldByName("TIPO").AsInteger = 2 
  EDITARVISAO.Enabled = (CurrentQuery.FieldByName("LINK").AsInteger = 2) Or (CurrentQuery.FieldByName("LINK").AsInteger = 3) Or (CurrentQuery.FieldByName("LINK").AsInteger = 8) 
 
  If (Not IsServerDeveloper) Then 
    EhDeSistema = (Mid(CurrentQuery.FieldByName("CODIGO").AsString, 1, 2) <> "K_") And (MenuDeSistema(CurrentQuery.FieldByName("GRUPO").AsInteger)) 
 
    ORDEM.ReadOnly = EhDeSistema 
    CODIGO.ReadOnly = EhDeSistema 
    NOME.ReadOnly = EhDeSistema 
    DESCRICAO.ReadOnly = EhDeSistema 
    ICONE.ReadOnly = EhDeSistema 
    LINK.ReadOnly = EhDeSistema 
    LINKGRUPO.ReadOnly = EhDeSistema 
    LINKVISAO.ReadOnly = EhDeSistema 
    LINKFECHADO.ReadOnly = EhDeSistema 
    LINKSQLESPECIAL.ReadOnly = EhDeSistema 
    PODEINSERIR.ReadOnly = EhDeSistema 
    PODEALTERAR.ReadOnly = EhDeSistema 
    PODEEXCLUIR.ReadOnly = EhDeSistema 
    LINKMODO.ReadOnly = EhDeSistema 
    LINKCONTEUDO.ReadOnly = EhDeSistema 
    LINKURL.ReadOnly = EhDeSistema 
    LINKTABVIRTUAL.ReadOnly = EhDeSistema 
    LINKRELATORIO.ReadOnly = EhDeSistema 
    LINKFILTRO.ReadOnly = EhDeSistema 
    LINKQUESTIONARIO.ReadOnly = EhDeSistema 
    TIPOSESPECIAIS.ReadOnly = EhDeSistema 
  End If 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  If CurrentQuery.FieldByName("LINK").AsInteger <> 1 Then 
    CurrentQuery.FieldByName("LINKGRUPO").Clear 
  End If 
  If CurrentQuery.FieldByName("LINK").AsInteger <> 2 And CurrentQuery.FieldByName("LINK").AsInteger <> 3 Then 
    CurrentQuery.FieldByName("LINKVISAO").Clear 
  End If 
  If CurrentQuery.FieldByName("LINK").AsInteger <> 4 Then 
    CurrentQuery.FieldByName("LINKCONTEUDO").Clear 
  End If 
  If CurrentQuery.FieldByName("LINK").AsInteger <> 8 Then 
    CurrentQuery.FieldByName("LINKTABVIRTUAL").Clear 
  End If 
  If CurrentQuery.FieldByName("LINK").AsInteger <> 9 Then 
    CurrentQuery.FieldByName("LINKRELATORIO").Clear 
    CurrentQuery.FieldByName("LINKFILTRO").Clear 
  End If 
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
 
Public Sub TABLE_OnDeleteBtnClick(CanContinue As Boolean) 
  If (Not IsServerDeveloper) Then 
    EhDeSistema = (Mid(CurrentQuery.FieldByName("CODIGO").AsString, 1, 2) <> "K_") And (MenuDeSistema(CurrentQuery.FieldByName("GRUPO").AsInteger)) 
    If (EhDeSistema) Then 
      CanContinue = False 
      MsgBox("Item de menu não pode ser excluído porque o menu é de sistema") 
    End If 
  End If 
End Sub 
 
Public Sub TABLE_UpdateRequired() 
  CODIGO_OnExit 
End Sub 
 
Public Function MenuDeSistema(aMenuHandle As Integer) As Boolean 
Dim QWork As Object 
  Set QWork = NewQuery 
  QWork.Text = "SELECT SISTEMA FROM W_GRUPOS WHERE HANDLE = :GRUPO" 
  QWork.ParamByName("GRUPO").AsInteger = aMenuHandle 
  QWork.Active = True 
  MenuDeSistema = QWork.FieldByName("SISTEMA").AsBoolean 
  QWork.Active = False 
  Set QWork = Nothing 
End Function 
