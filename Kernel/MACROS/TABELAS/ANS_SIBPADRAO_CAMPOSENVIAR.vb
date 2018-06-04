'HASH: 9191E63011199A1D5AFE224318D771A0
 
'Macro: ANS_SIBPADRAO_CAMPOSENVIAR
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If GerouXML Then
    bsShowMessage("Não é possível excluir um novo campo pois o arquivo XML já foi gerado.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If GerouXML Then
    bsShowMessage("Não é possível inserir um novo campo pois o arquivo XML já foi gerado.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Function GerouXML As Boolean

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(1) QTDE                    ")
  SQL.Add("  FROM ANS_SIBPADRAO_CAMPOSENVIAR       ")
  SQL.Add(" WHERE BENEFICIARIOCADASTRO = :CADASTRO ")
  SQL.Add("   AND BENEFICIARIOXML IS NOT NULL      ")

  SQL.ParamByName("CADASTRO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIOCADASTRO").AsInteger

  SQL.Active = True

  If SQL.FieldByName("QTDE").AsInteger = 0 Then
	GerouXML = False
  Else
	GerouXML = True
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If GerouXML Then
    bsShowMessage("Não é possível incluir um novo campo pois o arquivo XML já foi gerado.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
