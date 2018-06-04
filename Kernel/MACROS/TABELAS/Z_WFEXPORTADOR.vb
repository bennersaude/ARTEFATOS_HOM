'HASH: 29B21EB6B155A750BBAD54D6FFE9C64E
 
 
Public Sub CAMINHO_OnBtnClick() 
	Dim folder As String 
	Dim layer As String 
	folder = selectFolder("Selecione uma pasta", CurrentQuery.FieldByName("CAMINHO").AsString) 
	layer = "" 
	If (folder <> "") Then 
 
		If (isCustomSystem) Then 
			layer = ".40" 
		End If 
 
		folder = folder & "workflow" & layer & ".wfl" 
		CurrentQuery.FieldByName("CAMINHO").AsString = folder 
	End If 
End Sub 
 
Public Sub CAMINHOBUILDER_OnBtnClick() 
	CurrentQuery.FieldByName("CAMINHOBUILDER").AsString = OpenDialog 
End Sub 
 
Public Sub EXPORTARJORNADAS_OnChange() 
	CurrentQuery.UpdateRecord() 
 
	If CurrentQuery.FieldByName("EXPORTARJORNADAS").AsBoolean Then 
		JORNADAS.ReadOnly = False 
	Else 
		JORNADAS.ReadOnly = True 
	End If 
End Sub 
 
Public Sub EXPORTARMODELOS_OnChange() 
	CurrentQuery.UpdateRecord() 
 
	If CurrentQuery.FieldByName("EXPORTARMODELOS").AsBoolean Then 
		MODELOS.ReadOnly = False 
 
		CAMINHOBUILDER.ReadOnly = False 
		SENHABUILDER.ReadOnly = False 
	Else 
		MODELOS.ReadOnly = True 
 
		CAMINHOBUILDER.ReadOnly = True 
		SENHABUILDER.ReadOnly = True 
	End If 
End Sub 
 
Public Sub EXPORTARPROCESSOS_OnChange() 
	CurrentQuery.UpdateRecord() 
 
	If CurrentQuery.FieldByName("EXPORTARPROCESSOS").AsBoolean Then 
		PROCESSOS.ReadOnly = False 
 
	Else 
		PROCESSOS.ReadOnly = True 
	End If 
End Sub 
 
Public Sub EXPORTARSISTEMAS_OnChange() 
	CurrentQuery.UpdateRecord() 
 
	If CurrentQuery.FieldByName("EXPORTARSISTEMAS").AsBoolean Then 
		SISTEMAS.ReadOnly = False 
	Else 
		SISTEMAS.ReadOnly = True 
	End If 
End Sub 
 
Public Sub TABLE_AfterScroll() 
	If NodePath = "ADM_WORKFLOW|PROCESSOS" Then 
		If CurrentQuery.InInsertion Then 
			CurrentQuery.FieldByName("EXPORTARPROCESSOS").AsBoolean = True 
		End If 
		PROCESSOS.ReadOnly = False 
	Else 
		PROCESSOS.ReadOnly = True 
	End If 
 
	If NodePath = "ADM_WORKFLOW|MODELOS" Then 
		If CurrentQuery.InInsertion Then 
			CurrentQuery.FieldByName("EXPORTARMODELOS").AsBoolean = True 
		End If 
		MODELOS.ReadOnly = False 
 
		CAMINHOBUILDER.ReadOnly = False 
		SENHABUILDER.ReadOnly = False 
	Else 
		MODELOS.ReadOnly = True 
 
		CAMINHOBUILDER.ReadOnly = False 
		SENHABUILDER.ReadOnly = False 
	End If 
 
	SISTEMAS.ReadOnly = True 
	JORNADAS.ReadOnly = True 
End Sub 
 
Public Function isCustomSystem() 
 
	Dim sSQL As BPesquisa 
	Set sSQL = NewQuery 
 
	sSQL.Text = "SELECT CUSTOMSYSTEM FROM Z_SISTEMA" 
	sSQL.Active = True 
 
	isCustomSystem = sSQL.FieldByName("CUSTOMSYSTEM").AsString = "S" 
 
	sSQL.Active = False 
	Set sSQL = Nothing 
 
 
End Function 
