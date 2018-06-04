'HASH: D2A59280E0108B3AE39D06A901FBFE43
Option Explicit 
 
Public Sub BOTAOGERAR_OnClick() 
  Dim obj As Object, b As Boolean 
  Set obj = CreateBennerObject("CSCube.Cubos") 
 
  b = obj.Gerar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set obj = Nothing 
 
  CurrentQuery.Active = False 
  CurrentQuery.Active = True 
 
  If b Then 
    BOTAOVISUALIZAR_OnClick 
  End If 
 
End Sub 
 
Public Sub BOTAOSALVAR_OnClick() 
Dim obj As Object 
Set obj = CreateBennerObject("CSCubeForms.Cubos") 
obj.Salvar(CurrentSystem,CurrentQuery.FieldByName("CUBO").AsInteger) 
Set obj = Nothing 
End Sub 
 
Public Sub BOTAOVISUALIZAR_OnClick() 
Dim obj As Object 
Set obj = CreateBennerObject("CSCubeForms.Cubos") 
obj.Visualizar(CurrentSystem,CurrentQuery.FieldByName("CUBO").AsInteger) 
Set obj = Nothing 
End Sub 
 
Public Sub TABLE_AfterScroll() 
Dim q As Object 
If CurrentQuery.FieldByName("CUBO").IsNull Then 
  BOTAOVISUALIZAR.Enabled = False 
Else 
Set q = NewQuery 
q.Add("SELECT COUNT(HANDLE) NRECS FROM Z_CUBOS WHERE HANDLE = :CUBO AND NOT(DATAULTIMAGERACAO IS NULL)") 
q.ParamByName("CUBO").AsInteger = CurrentQuery.FieldByName("CUBO").AsInteger 
q.Active = True 
BOTAOVISUALIZAR.Enabled = (q.FieldByName("NRECS").AsInteger > 0) 
q.Active = False 
Set q = Nothing 
End If 
BOTAOSALVAR.Enabled = BOTAOVISUALIZAR.Enabled 
End Sub 
 
  Public Sub TABLE_BeforePost(CanContinue As Boolean) 
TABLE_AfterScroll 
End Sub 
 
