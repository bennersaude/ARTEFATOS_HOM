'HASH: FA62F47E0DA4364BF3FFDBF5730CF646
 
 
Public Sub EDITAR_OnClick() 
	Dim WebApiObj As Object 
	Dim WebApiPortNumber As Integer 
	Dim DataSourceEditorProxy As Object 
 
 
	Set WebApiObj = CreateBennerObject("Benner.Tecnologia.ManagedInterop.WebApiHostProxy") 
	WebApiPortNumber = WebApiObj.Start() 
 
	Set DataSourceEditorProxy = CreateBennerObject("Benner.Tecnologia.ManagedInterop.DataSourceEditorProxy") 
	DataSourceEditorProxy.Start(WebApiPortNumber, CurrentQuery.FieldByName("IDENTIFICACAO").AsString) 
 
	Set WebApiObj = Nothing 
	Set DataSourceEditorProxy = Nothing 
End Sub 
 
Public Sub PREVER_OnClick() 
	Dim ReportObj As Object 
 
	Set ReportObj = CreateBennerObject("Benner.Tecnologia.ManagedInterop.StimulsoftReportProxy") 
	ReportObj.CreateDynamicStimulsoftReport(CurrentQuery.FieldByName("HANDLE").AsInteger) 
	Set ReportObj = Nothing 
End Sub 
