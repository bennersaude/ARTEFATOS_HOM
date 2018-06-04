'HASH: F455299D274B03A761DA85BED26A4D5A
 
 
Public Sub EXPORTTEMPLATE_OnClick() 
 
Dim ws As WorkflowWebService 
Set ws = NewWorkflowWebService 
 
ws.Start("teste") 
Set ws = Nothing 
Exit Sub 
 
 
Dim o As Object 
  Set o = CreateBennerObject("csworkflowforms.modulo") 
  o.ShowWorkFlow(CurrentSystem) 
  Set o = Nothing 
Exit Sub 
 
  Dim obj As Object 
    Set obj = CreateBennerObject("Benner.Tecnologia.Workflow.Application.TemplateExporterForm") 
    obj.Exec(CurrentSystem, False) 
    Set obj = Nothing 
End Sub 
 
Public Sub IMPORTTEMPLATE_OnClick() 
  Dim obj As Object 
  Set obj = CreateBennerObject("Benner.Tecnologia.Workflow.Application.TemplateExporterForm") 
    obj.Exec(CurrentSystem, True) 
    Set obj = Nothing 
End Sub 
