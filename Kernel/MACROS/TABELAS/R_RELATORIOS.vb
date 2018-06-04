'HASH: D64F259E8CF631B9CE9A682BF29F42BE
 
 
Public Sub EDITAR_OnClick() 
	Dim ReportObj As Object 
 
	Set ReportObj = CreateBennerObject("Benner.Tecnologia.ManagedInterop.StimulsoftReportProxy") 
	ReportObj.LoadReport(CurrentQuery.FieldByName("HANDLE").AsInteger) 
	ReportObj.Edit 
 
 
	Set ReportObj = Nothing 
End Sub 
 
Public Sub EXPORTAR_OnClick() 
	Dim ExpObj As Object 
 
	Set ExpObj = CreateBennerObject("CS.RelExportar") 
	ExpObj.Exec 
	Set ExpObj = Nothing 
End Sub 
 
Public Sub EXCLUIR_OnClick() 
	Dim ExcObj As Object 
	Set ExcObj = CreateBennerObject("CS.RelExcluir") 
	ExcObj.Exec 
	Set ExcObj = Nothing 
End Sub 
 
Public Sub PREVER_OnClick() 
 
	Dim rep As CSReportPrinter 
	Dim recentItem As Object 
 
	On Error GoTo ProcessaErro 
 
		Set rep = NewReport(CurrentQuery.FieldByName("HANDLE").AsInteger) 
 
		rep.Preview 
 
		Set rep = Nothing 
 
		Set recentItem = CreateManagedObject("Benner.Tecnologia.Runner.UserTools", "Benner.Tecnologia.Runner.UserTools.ReportRibbonRecentItemProxy") 
 
		recentItem.AddReportToRibbonRecentItems(CurrentQuery.FieldByName("HANDLE").AsInteger, "") 
 
		Set recentItem = Nothing 
 
		Exit Sub 
 
ProcessaErro: 
 
	Err.Raise vbsUserException, "Stimulsoft Preview", Err.Description 
 
End Sub 
 
Public Sub TABLE_AfterScroll() 
	If (Not WebMode) Then 
		PREVER.Enabled = (CurrentQuery.State = 1) 
		EDITAR.Enabled = (CurrentQuery.State = 1) 
		EXPORTAR.Enabled = (CurrentQuery.State = 1) 
	End If 
End Sub 
 
Public Sub TABLE_NewRecord() 
	If (NodeInternalCode = 1002) Then 
		CurrentQuery.FieldByName("TIPO").AsInteger = 2 
	Else 
		CurrentQuery.FieldByName("TIPO").AsInteger = 1 
	End If 
End Sub 
 
Public Sub TIRARFONTES_OnClick() 
	Dim Sql, Sql2 
 
	If MsgBox("Deseja limpar todas a fontes desse relatório? ", vbOkCancel) = vbCancel Then 
		Exit Sub 
	End If 
 
	If InTransaction Then 
		Rollback 
	End If 
	StartTransaction 
 
	On Error GoTo ProcessaErro 
 
	Set Sql = NewQuery 
	Set Sql2 = NewQuery 
   	Sql.Add "UPDATE R_RELATORIOS SET FONTE = NULL, FONTELEGENDAS = NULL WHERE HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString 
   	Sql.ExecSQL 
 
   	Sql.Active = False 
   	Sql.Clear 
   	Sql.Add "UPDATE R_DETALHES SET FONTE = NULL WHERE RELATORIO = " + CurrentQuery.FieldByName("HANDLE").AsString 
	Sql.ExecSQL 
 
	Sql.Active = False 
   	Sql.Clear 
   	Sql.Add "SELECT HANDLE FROM R_DETALHES WHERE RELATORIO = " + CurrentQuery.FieldByName("HANDLE").AsString 
	Sql.Active = True 
	While Not Sql.EOF 
   		Sql2.Active = False 
   		Sql2.Clear 
   		Sql2.Add "UPDATE R_DETALHEDETALHES SET FONTELEGENDAS = NULL WHERE DETALHE = " + Sql.FieldByName("HANDLE").AsString 
		Sql2.ExecSQL 
   		Sql2.Active = False 
   		Sql2.Clear 
   		Sql2.Add "UPDATE R_DETALHECAMPOS SET FONTE = NULL WHERE DETALHE = " + Sql.FieldByName("HANDLE").AsString 
		Sql2.ExecSQL 
		Sql.Next 
	Wend 
 
	Set Sql = Nothing 
	Set Sql2 = Nothing 
 
	Commit 
	Exit Sub 
 
ProcessaErro: 
	Rollback 
End Sub 
