'HASH: A4F09AD77ACAA00502C3D830CBD3D488
Option Explicit
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim vSQL As Object

	Set vSQL = NewQuery
	vSQL.Clear
	vSQL.Add("SELECT CODNEGACAOEXTERNO")
	vSQL.Add("  FROM AEX_NEGACAO")
	vSQL.Add(" WHERE CODNEGACAOSISTEMA = :NEGACAO")
	vSQL.ParamByName("NEGACAO").AsInteger = CurrentQuery.FieldByName("MOTIVONEGACAOSISTEMA").AsInteger
	vSQL.Active = True

	If (vSQL.EOF)  Then
		CanContinue = False
		bsShowMessage("O motivo de negação selecionado não foi encontrado na tabela de conversão de negações da empresa.", "E")
	Else
		If vSQL.FieldByName("CODNEGACAOEXTERNO").AsInteger > 0 Then
			CurrentQuery.FieldByName("MOTIVONEGACAO").AsInteger = vSQL.FieldByName("CODNEGACAOEXTERNO").AsInteger
		Else
			CanContinue = False
			bsShowMessage("O motivo de negação foi encontrado na tabela de conversão mas o código não é válido.", "E")
		End If
	End If
	Set vSQL = Nothing
End Sub
