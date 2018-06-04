'HASH: 90DB5AEA8201BD0B8D7F7734EF080C87

Option Explicit

Dim SQL As Object
Dim VerificaStatus As Integer


Public Sub TABLE_AfterScroll()
	Set SQL = NewQuery
	SQL.Clear

	SQL.Add("SELECT A.SITUACAOGERACAO   										 ")
  	SQL.Add("  FROM SAM_ROTINASUSPENSAO A										 ")
    SQL.Add("  Join SAM_ROTINASUSPENSAOFAMILIA B On A.Handle = B.ROTINASUSPENSAO ")
 	SQL.Add(" WHERE B.ROTINASUSPENSAO = :ROTINASUSPENSAO                         ")

	SQL.ParamByName("ROTINASUSPENSAO").Value = CurrentQuery.FieldByName("ROTINASUSPENSAO").AsInteger

	SQL.Active = True

	If (SQL.FieldByName("SITUACAOGERACAO").Value <> 1) Then
		OBSERVACAO.ReadOnly = True
	Else
		OBSERVACAO.ReadOnly = False
	End If

	Set SQL = Nothing
End Sub
