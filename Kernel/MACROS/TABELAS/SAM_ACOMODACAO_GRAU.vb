'HASH: 55F93CB5C8ADF09A7CF77E259FFD8E65
 

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  'Keila -SMS 7374
  Dim vProcura As Object
  Dim handlex As Long
  ShowPopup =False
  Set vProcura =CreateBennerObject("Procura.Procurar")
  handlex =vProcura.Exec(CurrentSystem,"SAM_GRAU","GRAU|DESCRICAO",2,"Código|Descrição","HANDLE > 0 AND TIPOGRAU IN (SELECT HANDLE FROM SAM_TIPOGRAU WHERE CLASSIFICACAO = '3')","Procura por Grau",True,"")
  If handlex <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value =handlex
  End If
  Set vProcura =Nothing

End Sub

Public Sub TABLE_AfterPost()
If CurrentQuery.FieldByName("GRAUPADROACOMODACAO").AsString ="S" Then
Dim SQL As Object
Set SQL =NewQuery
SQL.Add("UPDATE SAM_ACOMODACAO_GRAU SET GRAUPADROACOMODACAO = 'N' WHERE ACOMODACAO = :A AND HANDLE <> :H")
SQL.ParamByName("A").Value =RecordHandleOfTable("SAM_ACOMODACAO")
SQL.ParamByName("H").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ExecSQL
Set SQL =Nothing
End If
End Sub

Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    GRAU.WebLocalWhere = " A.TIPOGRAU IN (SELECT HANDLE FROM SAM_TIPOGRAU WHERE CLASSIFICACAO = '3') "
  End If
End Sub
