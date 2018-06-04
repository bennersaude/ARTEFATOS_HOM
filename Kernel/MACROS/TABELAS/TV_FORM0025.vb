'HASH: FB39DC3409A08ACE827EF48D10C0708D
 

Option Explicit

'--------------------------------------------------------------------------------------------------------------------------
'  SOMENTE USAR A PARTIR DA SAM_AUTORIZ_EVENTOSOLICIT ---------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------


Public Sub AlterarEquipeVia
	Dim sql As BPesquisa
	Set sql = NewQuery
	sql.Add("UPDATE SAM_AUTORIZ_EVENTOSOLICIT SET EQUIPE=:E, VIAACESSO=:V WHERE HANDLE=:H")
	sql.ParamByName("E").Value = CurrentQuery.FieldByName("EQUIPE").AsInteger
	sql.ParamByName("v").Value = CurrentQuery.FieldByName("VIAACESSO").AsInteger
	sql.ParamByName("H").Value = RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT")
	sql.ExecSQL
	Set sql=Nothing
End Sub


Public Sub TABLE_AfterPost()
	AlterarEquipeVia
End Sub

