'HASH: 0276C61ECDC6D5E3C1618B07FA3C1BB7
 
 
Public Sub BOTAOEDITARCAMPOS_OnClick() 
Dim obj As Object 
  If CurrentQuery.State <> 1 Then Exit Sub 
  Set obj = CreateBennerObject("CSCubeForms.CSDesktop") 
  obj.EditDataPacketSQL(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set obj = Nothing 
  Rem Apenas nesta situação pode ser feita a operação abaixo 
  CurrentQuery.Active = False 
  CurrentQuery.Active = True 
End Sub 
 
Public Sub BOTAOEDITARFILTRO_OnClick() 
Dim s As String, obj As Object 
If CurrentQuery.State <> 1 Then Exit Sub 
Set obj = CreateBennerObject("CSCubeForms.CSDesktop") 
obj.EditarFiltro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("FILTROPADRAO").AsInteger) 
Set obj = Nothing 
CurrentQuery.Active = False 
CurrentQuery.Active = True 
TABLE_AfterScroll 
End Sub 
 
Public Sub BOTAOEDITARPLUGIN_OnClick() 
Dim s As String, obj As Object, q As Object 
Set obj = CreateBennerObject("Benner.Tecnologia.Desktop.Forms.DataPacketPluginEditor") 
On Error GoTo PluginError 
If obj.EditPlugin(CurrentQuery.FieldByName("COMANDOSQL").AsString) Then 
  If CurrentQuery.State = 1 Then 
    Set q = NewQuery 
    q.Add("UPDATE Z_DATAPACKETS SET COMANDOSQL = :XML, DESCRICAOPLUGIN = :DESCRICAO, USUARIOALTEROU = :USUARIO, DATAALTERACAO = :DATA WHERE HANDLE = :HANDLE") 
    q.ParamByName("XML").AsMemo = obj.xml 
    q.ParamByName("DESCRICAO").AsString = obj.Description 
    q.ParamByName("USUARIO").AsInteger = CurrentUser 
    q.ParamByName("DATA").AsDateTime = ServerNow 
    q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
    StartTransaction 
    On Error GoTo erro 
    q.ExecSQL 
    Commit 
    GoTo sai 
erro: 
    Rollback 
    Err.Raise(10, "Não foi possível gravar alterações.") 
sai: 
    CurrentQuery.Active = False 
    CurrentQuery.Active = True 
  Else 
    CurrentQuery.FieldByName("COMANDOSQL").AsString	= obj.xml 
    CurrentQuery.FieldByName("DESCRICAOPLUGIN").AsString = obj.Description 
  End If 
End If 
Set obj = Nothing 
Exit Sub 
 
PluginError: 
WriteBDebugMessage("Erro do plug-in: " + Err.Description) 
MsgBox "Ocorreu um erro ao editar o plug-in." 
End Sub 
 
Public Sub BOTAOEXCLUIRFILTRO_OnClick() 
Dim q As Object 
If (CurrentQuery.State <> 1) Or (CurrentQuery.FieldByName("FILTROPADRAO").IsNull) Then Exit Sub 
If MsgBox("Confirma exclusão do filtro ?",vbYesNo,"Confirmação") = vbYes Then 
Set q = NewQuery 
StartTransaction 
q.Add("DELETE FROM Z_FILTROCONDICOES WHERE FILTRO = "+CStr(CurrentQuery.FieldByName("FILTROPADRAO").AsInteger)) 
q.ExecSQL 
q.Clear 
q.Add("UPDATE Z_DATAPACKETS SET FILTROPADRAO = NULL WHERE HANDLE = "+CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)) 
q.ExecSQL 
q.Clear 
q.Add("DELETE FROM Z_FILTROS WHERE HANDLE = "+CStr(CurrentQuery.FieldByName("FILTROPADRAO").AsInteger)) 
q.ExecSQL 
Set q = Nothing 
Commit 
CurrentQuery.Active = False 
CurrentQuery.Active = True 
TABLE_AfterScroll 
End If 
End Sub 
 
Public Sub BOTAOGERAR_OnClick() 
Dim obj As Object 
  Set obj = CreateBennerObject("CSCUBE.DASHBOARDS") 
  obj.BuildDataPacket(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set obj = Nothing 
  CurrentQuery.Active = False 
  CurrentQuery.Active = True 
End Sub 
 
Public Sub BOTAOSALVAR_OnClick() 
Dim obj As Object 
  Set obj = CreateBennerObject("CSCUBEFORMS.CSDESKTOP") 
  obj.SaveDataPacket(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set obj = Nothing 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  CONTEUDO.ReadOnly = CurrentQuery.FieldByName("TIPO").AsInteger <> 2 
  BOTAOEXCLUIRFILTRO.Enabled = (Not CurrentQuery.FieldByName("FILTROPADRAO").IsNull) 
End Sub 
 
Public Sub TIPO_OnChange() 
  CurrentQuery.UpdateRecord 
  CONTEUDO.ReadOnly = CurrentQuery.FieldByName("TIPO").AsInteger <> 2 
End Sub 
