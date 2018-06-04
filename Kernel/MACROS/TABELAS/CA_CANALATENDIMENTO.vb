'HASH: 4077126B3413DEC8C3F4915337327CB1
'MACRO = CA_CANALATENDIMENTO

'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim canalAtendimento As String
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DESCRICAO FROM CA_CANALATENDIMENTO WHERE PADRAO = 'S'")
  SQL.Add("    AND HANDLE <> :HANDLE_CANAL")

  SQL.ParamByName("HANDLE_CANAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  SQL.Active = True

  If (Not SQL.EOF) And (CurrentQuery.FieldByName("PADRAO").AsString = "S") Then
	bsShowMessage("'"+ SQL.FieldByName("DESCRICAO").AsString + "'" + " já está cadastrado como Canal de Atendimento padrão!", "E")
	CanContinue = False
  End If

  Set SQL = Nothing

End Sub
