'HASH: 85F30FF15EF8B9C225A8ACA811D92C98
'MACRO: MS_PARAMETROGERAL


Public Sub CODIGORELATORIOASO_OnBtnClick()
 'verificar o nome do campo em cada tabela
  Dim ProcuraDLL As Object
  Dim Handlexx As Long
  Set ProcuraDLL = CreateBennerObject("Procura.Procurar")
  Handlexx = ProcuraDLL.Exec(CurrentSystem,"R_RELATORIOS","CODIGO|NOME",1,"Código|Nome","","Relatórios do sistema",True,"")

  If Handlexx <>0 Then
    Dim SQL As Object
    Set SQL =NewQuery
    SQL.Add("SELECT CODIGO, NOME FROM R_RELATORIOS WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value =Handlexx
    SQL.Active =True

    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If

    CurrentQuery.FieldByName("CODIGORELATORIOASO").Value = SQL.FieldByName("CODIGO").AsString  'verificar os nomes dos campos
    CurrentQuery.FieldByName("NOMERELATORIOASO").Value   = SQL.FieldByName("NOME").AsString    'verificar os nomes dos campos
  End If

  Set ProcuraDLL = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  CODIGORELATORIOASO.ReadOnly = True
End Sub

'SMS 87108 - Ricardo Rocha - 05/10/2007 - OnCommandClick para WEB no lugar de OnClick
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "CODIGORELATORIOASO") Then
		CODIGORELATORIOASO_OnBtnClick
	End If
End Sub
