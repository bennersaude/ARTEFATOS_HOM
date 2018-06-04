'HASH: 28FDEB02604AF76E2FF0259D451F7712
'Macro: R_GRUPORELATORIOS
'#Uses "*EnviarRelatorioEmail"

Public Sub IMPRIMIRRELATORIO_OnClick()
  EnviarRelatorioEmail(CurrentQuery.FieldByName("RELATORIO").AsInteger)
End Sub

Public Sub RELATORIO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim QueryUpdateModulo As Object

  Set QueryUpdateModulo = NewQuery

  vColunas = "NOME"

  vCampos = "Descrição"
  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")
  interface.seleciona(CurrentSystem, "R_RELATORIOS", vColunas, vCampos, "R_GRUPORELATORIOS", "GRUPO", RecordHandleOfTable("R_GRUPOSRELATORIOS"), "RELATORIO", "Seleciona Relatório")
  Set interface = Nothing

  If Not InTransaction Then StartTransaction

  QueryUpdateModulo.Add("UPDATE R_GRUPORELATORIOS  ")
  QueryUpdateModulo.Add("   SET MODULO = :P_MODULO ")
  QueryUpdateModulo.Add(" WHERE GRUPO  = :P_GRUPO  ")
  QueryUpdateModulo.ParamByName("P_GRUPO").AsInteger = RecordHandleOfTable("R_GRUPOSRELATORIOS")
  QueryUpdateModulo.ParamByName("P_MODULO").AsInteger = CurrentModule
  QueryUpdateModulo.ExecSQL

  Commit
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT CODIGO FROM R_RELATORIOS WHERE HANDLE = :RELATORIO")
  SQL.ParamByName("RELATORIO").Value = CurrentQuery.FieldByName("RELATORIO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("CODIGO").AsString = "DEM-RC8" Then
    IMPRIMIRRELATORIO.Visible = True
  Else
    IMPRIMIRRELATORIO.Visible = False
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "IMPRIMIRRELATORIO" Then
		IMPRIMIRRELATORIO_OnClick
	End If
End Sub
