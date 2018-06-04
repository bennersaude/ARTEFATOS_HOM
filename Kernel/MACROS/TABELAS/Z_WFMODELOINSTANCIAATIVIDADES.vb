'HASH: 55790568DE8A8D2D9A301EE2ED960DB5
Option Explicit
 
Public Sub TABLE_AfterScroll()
  Dim i As Integer
 
  Dim nomeTabela As String 
 
  Dim Q As BPesquisa 
  Set Q = NewQuery 
  Q.Add("SELECT Z_TABELAS.NOME FROM Z_WFMODELOINSTANCIAS, Z_WFMODELOS, Z_TABELAS WHERE (Z_WFMODELOINSTANCIAS.HANDLE = :MODELOINSTANCIA And Z_WFMODELOINSTANCIAS.MODELO = Z_WFMODELOS.HANDLE) AND (Z_TABELAS.HANDLE = Z_WFMODELOS.TABELA OR (Z_TABELAS.HANDLE = (SELECT X1.TABELA FROM Z_WFMODELOINSTANCIAS X3, Z_WFMODELOS X2, W_VISOES X1 WHERE X3.HANDLE = :MODELOINSTANCIA AND X3.MODELO = X2.HANDLE AND X1.HANDLE = X2.VISAO)))") 
  Q.ParamByName("MODELOINSTANCIA").AsInteger = CurrentQuery.FieldByName("MODELOINSTANCIA").AsInteger 
  Q.Active = True 
  nomeTabela = Q.FieldByName("NOME").AsString 
  Q.Active = False 
  TABELADADORELEVANTE.Text = "Tabela: " + nomeTabela 
 
  If (Not CurrentQuery.FieldByName("HANDLEDADORELEVANTE").IsNull ) Then 
    Q.Clear 
    Q.Add("SELECT * FROM " + nomeTabela + " WHERE HANDLE="+ CurrentQuery.FieldByName("HANDLEDADORELEVANTE").AsString) 
    Q.Active = True 
    DADOSRELEVANTES.Text = "" 
  If (Not Q.EOF) Then 
    For i=0 To Q.FieldCount-1 
      DADOSRELEVANTES.Text = DADOSRELEVANTES.Text + Q.Fields(i).FieldName + ": " + Q.Fields(i).AsString + Chr(13) + Chr(10) 
    Next i 
  End If 
    Q.Active = False 
  End If 
 
  Set Q = Nothing 
End Sub 
