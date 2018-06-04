'HASH: 1F314663B7CDC2A0A31DF457007AA2CD
Option Explicit 
 
Public Sub TABLE_AfterPost() 
Dim filtro As Boolean 
Dim q As BPesquisa 
Set q = NewQuery 
q.Add("UPDATE Z_EMAILS SET STATUS = 2 WHERE STATUS = 6") 
If VisibleMode Then 
	If NodeInternalCode = 6 Then 
  	q.Add(" AND USUARIO = " + CStr(CurrentUser)) 
	End If 
End If 
 
If Not CurrentQuery.FieldByName("INICIO").IsNull Then 
	q.Add(" AND DATAINCLUSAO >= :INICIO") 
End If 
 
If Not CurrentQuery.FieldByName("FIM").IsNull Then 
	q.Add(" AND DATAINCLUSAO <= :FIM") 
End If 
 
If Not CurrentQuery.FieldByName("USUARIO").IsNull Then 
	q.Add(" AND USUARIO = " + CStr(CurrentQuery.FieldByName("USUARIO").AsInteger)) 
End If 
 
If Not CurrentQuery.FieldByName("SELECAOESPECIAL").IsNull Then 
	q.Add(" AND (" + CurrentQuery.FieldByName("SELECAOESPECIAL").AsString + ")") 
End If 
 
If Not CurrentQuery.FieldByName("INICIO").IsNull Then 
	q.ParamByName("INICIO").AsDateTime = CurrentQuery.FieldByName("INICIO").AsDateTime 
End If 
 
If Not CurrentQuery.FieldByName("FIM").IsNull Then 
	q.ParamByName("FIM").AsDateTime = CurrentQuery.FieldByName("FIM").AsDateTime 
End If 
 
q.ExecSQL 
 
Set q = Nothing 
 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
 
If CurrentQuery.FieldByName("INICIO").IsNull And CurrentQuery.FieldByName("FIM").IsNull And CurrentQuery.FieldByName("USUARIO").IsNull And CurrentQuery.FieldByName("SELECAOESPECIAL").IsNull Then 
	Err.Raise vbsUserException, vbEmpty, "Nenhum filtro foi informado." 
End If 
 
End Sub 
