'HASH: 9DB9E398F673F4FFA00992BE294312DD
 
Public Sub EDITOR_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Benner.Tecnologia.Workflow.Application.DesignerForm") 
    Obj.Exec(CurrentSystem, -1, False) 
    ' caso haja inserção de modelo(s) devemos atualizar os registros na árvore 
    If ( Obj.ShouldRefreshNodesWithTable ) Then 
      RefreshNodesWithTable ( "Z_WFMODELOS" ) 
    End If 
    Set Obj = Nothing 
End Sub 
 
Public Sub EXPORTTEMPLATE_OnClick() 
  Dim form As CSVirtualForm 
  Set form = NewVirtualForm 
  form.TableName = "Z_WFEXPORTADOR" 
  form.Caption = "Exportar artefatos do Workflow" 
  form.Show 
  Set form = Nothing 
End Sub 
 
Public Sub IMPORTTEMPLATE_OnClick() 
  Dim form As CSVirtualForm 
  Set form = NewVirtualForm 
  form.TableName = "Z_WFIMPORTADOR" 
  form.Caption = "Importar artefatos do Workflow" 
 
  Dim invalids As String 
 
  form.TransitoryVars("WORKFLOW.IMPORT.MODELSTOVALIDATE") = "" 
  form.Show 
  models = form.TransitoryVars("WORKFLOW.IMPORT.MODELSTOVALIDATE") 
  While models <> "" 
  	handle = ShortHint(models) 
	If ValidateModel(handle) = False Then 
      invalids = invalids & handle & "," 
	End If 
	models = LongHint(models) 
  Wend 
  Set form = Nothing 
 
  ' Identifica modelos que não puderam ser ativados, entao avisa o usuario 
  Dim msg As String 
  msg = "Os seguintes fluxos não puderam ser validados automaticamente devido a algumas inconsistências nas suas configurações. Será necessário efetuar a validação manualmente através do editor de fluxos." 
  If invalids <> "" Then 
	Dim q As BPesquisa 
	Set q = NewQuery 
	q.Text = "SELECT NOME,VERSAO FROM Z_WFMODELOS WHERE HANDLE IN (" & invalids & "-1)" 
	q.Active = True 
	While Not q.EOF 
		msg = msg & Chr(13) & Chr(10) & "- " & q.FieldByName("NOME").AsString & " - v." & q.FieldByName("VERSAO").AsInteger 
		q.Next 
	Wend 
	q.Active = False 
	Set q = Nothing 
	MsgBox(msg) 
  End If 
End Sub 
 
Public Function ValidateModel(modelHandle) As Boolean 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Benner.Tecnologia.Workflow.Application.TemplateExporter") 
  ValidateModel = Obj.ValidateModel(modelHandle, CurrentSystem) 
  Obj.UpdateModelImage(modelHandle, CurrentSystem) 
  Set Obj = Nothing 
End Function 
 
Public Sub PUBLISH_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishWFFieldsVisibility(CurrentSystem, -1) 
  Set Obj = Nothing 
End Sub 
