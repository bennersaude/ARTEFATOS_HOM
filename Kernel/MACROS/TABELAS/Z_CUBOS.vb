'HASH: 8CE9D42996FB2D0BF1E122AFED08CE58
Option Explicit 
Public Sub BOTAOEDITAR_OnClick() 
Dim obj As Object 
Set obj = CreateBennerObject("CSCubeForms.Cubos") 
obj.Editar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) 
Set obj = Nothing 
CurrentQuery.Active = False 
CurrentQuery.Active = True 
TABLE_AfterScroll 
End Sub 
Public Sub BOTAOEXCLUIRFILTRO_OnClick() 
Dim q As Object 
If (CurrentQuery.State <> 1) Or (CurrentQuery.FieldByName("FILTROPADRAO").IsNull) Then Exit Sub 
If MsgBox("Confirma exclusão do filtro ?",vbYesNo,"Confirmação") = vbYes Then 
Set q = NewQuery 
StartTransaction 
q.Add("DELETE FROM Z_FILTROCONDICOES WHERE FILTRO = "+CStr(CurrentQuery.FieldByName("FILTROPADRAO").AsInteger)) 
q.ExecSQL 
q.Clear 
q.Add("UPDATE Z_CUBOS SET FILTROPADRAO = NULL WHERE HANDLE = "+CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)) 
q.ExecSQL 
q.Clear 
q.Add("DELETE FROM Z_FILTROS WHERE HANDLE = "+CStr(CurrentQuery.FieldByName("FILTROPADRAO").AsInteger)) 
q.ExecSQL 
Set q = Nothing 
Commit 
CurrentQuery.Active = False 
CurrentQuery.Active = True 
TABLE_AfterScroll 
End If 
End Sub 
Public Sub BOTAOFILTRO_OnClick() 
Dim s As String, obj As Object 
If CurrentQuery.State <> 1 Then Exit Sub 
Set obj = CreateBennerObject("CSCubeForms.Cubos") 
obj.DefinirFiltro(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,CurrentQuery.FieldByName("FILTROPADRAO").AsInteger) 
Set obj = Nothing 
CurrentQuery.Active = False 
CurrentQuery.Active = True 
TABLE_AfterScroll 
End Sub 
Public Sub BOTAOGERAR_OnClick() 
Dim obj As Object, b As Boolean 
Set obj = CreateBennerObject("CSCube.Cubos") 
b = obj.Gerar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) 
Set obj = Nothing 
CurrentQuery.Active = False 
CurrentQuery.Active = True 
If b Then BOTAOVER_OnClick 
End Sub 
Public Sub BOTAOSALVAR_OnClick() 
Dim obj As Object 
Set obj = CreateBennerObject("CSCubeForms.Cubos") 
obj.Salvar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) 
Set obj = Nothing 
End Sub 
Public Sub BOTAOVER_OnClick() 
Dim obj As Object 
Set obj = CreateBennerObject("CSCubeForms.Cubos") 
obj.Visualizar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) 
Set obj = Nothing 
End Sub 
Public Sub TABELA_OnChange() 
CurrentQuery.FieldByName("CAMPOS").Clear 
End Sub 
Public Sub TABLE_AfterInsert() 
If Not LicensedTool("BI") Then 
CurrentQuery.FieldByName("NAOGERARVISAO").AsString = "N" 
CurrentQuery.FieldByName("MANTERTODASGERACOES").AsString = "N" 
End If 
End Sub 
 
Public Sub TABLE_AfterPost() 
TABLE_AfterScroll 
End Sub 
Public Sub TABLE_AfterScroll() 
Dim grant As Boolean 
grant = LicensedTool("BI") 
BOTAOGERAR.Enabled = (Not CurrentQuery.FieldByName("CAMPOS").IsNull) 
BOTAOVER.Enabled = (Not CurrentQuery.FieldByName("DATAULTIMAGERACAO").IsNull) Or (Not grant) 
BOTAOSALVAR.Enabled = BOTAOVER.Enabled And grant 
BOTAOEXCLUIRFILTRO.Enabled = (Not CurrentQuery.FieldByName("FILTROPADRAO").IsNull) 
MANTERTODASGERACOES.Visible = grant 
NAOGERARVISAO.Visible = grant 
If Not grant Then 
  BOTAOVER.Caption = "Gerar" 
End If 
End Sub 
 
Public Sub TABLE_BeforeDelete(CanContinue As Boolean) 
Dim obj As Object 
Set obj = CreateBennerObject("CSCUBE.CUBOS") 
obj.ExcluirCubo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
Set obj = Nothing 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
	ValidarCaracteresInvalidosCodigo 
	ValidarCodigoUnico 
End Sub 
 
Public Sub ValidarCaracteresInvalidosCodigo() 
	Dim validar As CSEntityCall 
 
	On Error GoTo ProcessaErro 
 
	Set validar = BusinessEntity.CreateCall("Benner.Tecnologia.Metadata.Entities.ZCubos, Benner.Tecnologia.Metadata", "ValidarCaracteresInvalidosCodigo") 
	validar.AddParameter(pdtString, CurrentQuery.FieldByName("CODIGO").AsString) 
	validar.Execute() 
	Set validar = Nothing 
 
	Exit Sub 
 
	ProcessaErro: 
	Set validar = Nothing 
	Err.Raise(vbsUserException, "Cadastro de cubos", Err.Description) 
 
End Sub 
 
 
Public Sub ValidarCodigoUnico() 
	Dim validar As CSEntityCall 
 
	On Error GoTo ProcessaErro 
 
	Set validar = BusinessEntity.CreateCall("Benner.Tecnologia.Metadata.Entities.ZCubos, Benner.Tecnologia.Metadata", "ValidarCodigoUnico") 
	validar.AddParameter(pdtString, CurrentQuery.FieldByName("CODIGO").AsString) 
	validar.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger) 
	validar.Execute() 
	Set validar = Nothing 
 
	Exit Sub 
 
	ProcessaErro: 
	Set validar = Nothing 
	Err.Raise(vbsUserException, "Cadastro de cubos", Err.Description) 
 
End Sub 
