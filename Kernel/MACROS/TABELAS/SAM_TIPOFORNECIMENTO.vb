'HASH: CB8356DE7470E088AA584A65AF9326AD
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim pSql As BPesquisa
Set pSql = NewQuery

'Inicio SMS  174752 - Leandro Manso - 12/12/2011
pSql.Add("SELECT COUNT(0) QUANTIDADE FROM SAM_TIPOFORNECIMENTO WHERE DEFINIDOPELODEPTOCOMPRAS = 'S'")
pSql.Active = True

If (pSql.FieldByName("QUANTIDADE").AsInteger > 0 And CurrentQuery.FieldByName("DEFINIDOPELODEPTOCOMPRAS").AsBoolean = True) Then
	CanContinue = False
 	bsShowMessage("Já existe registro com campo 'Definido pelo departamento de compras' marcado!", "I")
	Set pSql = Nothing
	Exit Sub
End If
Set pSql = Nothing
'Fim SMS 174752 - Leandro Manso - 12/12/2011
End Sub

