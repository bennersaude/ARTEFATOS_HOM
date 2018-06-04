'HASH: 94D1D370806384B87A1562450E255EDD
Option Explicit 
 
Public Sub Registros_OnClick() 
  Dim sql As Object 
  Set sql = NewQuery 
  sql.Add("SELECT NOME FROM Z_TABELAS WHERE HANDLE = " + CurrentQuery.FieldByName("TABELA").AsString) 
  sql.Active = True 
 
  If NodeInternalCode = 11 Then 
   Filter(sql.FieldByName("NOME").AsString + "*RESTRICOESPORREGISTRO*|" + CurrentQuery.FieldByName("HANDLE").AsString) 
  Else 
    Filter(sql.FieldByName("NOME").AsString + "*PERMISSAO*") 
  End If 
 
  Set sql = Nothing 
 
  TABLE_AfterScroll 
End Sub 
 
Public Sub TABLE_AfterInsert() 
 
  CurrentQuery.FieldByName("NOME").AsString = "Filtro de Permissão" 
  CurrentQuery.FieldByName("PERMISSAO").AsString = "S" 
  CurrentQuery.FieldByName("CATEGORIA").AsInteger = 2 
 
  If (VisibleMode) Then 
    Select Case NodeInternalCode 
 
      Case 10  ' *** Filtro de registros de grupo 
        CurrentQuery.FieldByName("GRUPO").AsInteger = RecordHandleOfTable("Z_GRUPOS") 
        CurrentQuery.FieldByName("USUARIO").Clear 
 
      Case 11  ' *** Restrições de alterações por registro 
        CurrentQuery.FieldByName("CATEGORIA").AsInteger = 4 
        CurrentQuery.FieldByName("NOME").AsString = "Restrições por registro" 
        CurrentQuery.FieldByName("GRUPO").AsInteger = RecordHandleOfTable("Z_GRUPOS") 
        CurrentQuery.FieldByName("USUARIO").Clear 
 
      Case Else  ' *** Filtro de registros de usuário 
        CurrentQuery.FieldByName("USUARIO").AsInteger = RecordHandleOfTable("Z_GRUPOUSUARIOS") 
        CurrentQuery.FieldByName("GRUPO").Clear 
 
    End Select 
  End If 
 
End Sub 
 
Public Sub TABLE_AfterScroll() 
 
  REGISTROS.Enabled =  Not CurrentQuery.FieldByName("TABELA").IsNull 
 
  If NodeInternalCode = 11 Then ' Não pode alterar a TABELA se já houver Restrições de alterações por registro para ela 
 
    TABELA.ReadOnly = (Not CurrentQuery.FieldByName("TABELA").IsNull) And HasDefinedConditions(CurrentQuery.FieldByName("HANDLE").AsInteger) 
  End If 
 
End Sub 
 
Public Sub TABLE_BeforeDelete(CanContinue As Boolean) 
  Dim q As Object 
  Set q = NewQuery 
  q.Add("DELETE FROM Z_FILTROCONDICOES WHERE FILTRO = "+CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)) 
  q.ExecSQL 
  Set q = Nothing 
End Sub 
 
Public Function HasDefinedConditions(FilterHandle As Integer) As Boolean 
 
  Dim q As BPesquisa 
  Set q = NewQuery 
 
  q.Text = "SELECT COUNT(*) REGISTROS FROM Z_FILTROCONDICOES WHERE FILTRO = :FILTRO" 
  q.ParamByName("FILTRO").AsInteger = FilterHandle 
  q.Active = True 
 
  HasDefinedConditions = q.FieldByName("REGISTROS").AsInteger > 0 
 
  Set q = Nothing 
 
End Function 
 
Public Sub TABLE_UpdateRequired() 
  If CurrentQuery.FieldByName("CATEGORIA").AsInteger <> 4 Then 
    OPERACAO.Required = False 
  End If 
End Sub 
