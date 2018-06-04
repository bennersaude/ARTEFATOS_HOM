'HASH: FA725D7297264CAD4A8607E44694DAFB
'#Uses "*bsShowMessage"

Option Explicit

Dim gTipoAnterior As String
Dim gDescricaoAnterior As String


Public Sub RELATORIO_OnBtnClick()
  If CurrentQuery.State = 1 Then
    bsShowMessage("Registro deve estar em edição ou inserção", "I")
  Else

  	If CurrentQuery.FieldByName("TIPO").AsString <> "1" Then
      Exit Sub
	End If

    Dim vHandleRelatorio As Long
    Dim Interface As Object
    Dim vCampos As String
    Dim vColunas As String
    Dim vCriterio As String
    Dim vTabela As String
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CODIGO, NOME FROM R_RELATORIOS WHERE HANDLE = :H")
    Set Interface = CreateBennerObject("Procura.Procurar")
    vColunas = "CODIGO|NOME"
    vCriterio = ""
    vCampos = "Código |Relatório"
    vTabela = "R_RELATORIOS"
    vHandleRelatorio = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, " CODIGO IN ('AUT010','AUT001','AUTOFAX','10001')", "Relatório", True, "")
    If vHandleRelatorio >0 Then
      SQL.Active = False
      SQL.ParamByName("H").Value = vHandleRelatorio
      SQL.Active = True
      CurrentQuery.FieldByName("RELATORIO").Value = SQL.FieldByName("CODIGO").AsString
    End If
    Set Interface = Nothing
    Set SQL = Nothing
  End If
End Sub

Public Sub TABLE_AfterScroll()
	If CurrentQuery.FieldByName("TIPO").AsString = "2" Then
		CLASSELIMINAR.Visible = True
	Else
		CLASSELIMINAR.Visible = False
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  gTipoAnterior = CurrentQuery.FieldByName("TIPO").AsString
  gDescricaoAnterior = CurrentQuery.FieldByName("DESCRICAO").AsString
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

CurrentQuery.UpdateRecord
	TABLE_AfterScroll

  If Len(CurrentQuery.FieldByName("OBSERVACAO").AsString) > 4000 Then
    bsShowMessage("Não é possível gravar uma observação com mais de 4000 caracteres.", "E")
    CanContinue = False
    Exit Sub
  End If

  If(gTipoAnterior = CurrentQuery.FieldByName("TIPO").AsString)And _
     (gDescricaoAnterior = CurrentQuery.FieldByName("DESCRICAO").AsString)Then
  Exit Sub
End If

Dim SQL As Object
Set SQL = NewQuery

SQL.Add("SELECT COUNT(HANDLE) QTD")
SQL.Add("FROM SAM_AUTORIZ_ANOTADM")
SQL.Add("WHERE ANOTACAO = :HANOTACAO")
SQL.ParamByName("HANOTACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True
If SQL.FieldByName("QTD").AsInteger >0 Then
  CanContinue = False
  bsShowMEssage("Esta anotação está sendo usada em " + SQL.FieldByName("QTD").AsString + " autorização(ões). Alteração não permitida", "E")
  Set SQL = Nothing
  Exit Sub
End If

SQL.Clear
SQL.Add("SELECT COUNT(HANDLE) QTD")
SQL.Add("FROM SAM_BENEFICIARIO_ANOTADM")
SQL.Add("WHERE ANOTACAO = :HANOTACAO")
SQL.ParamByName("HANOTACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True
If SQL.FieldByName("QTD").AsInteger >0 Then
  CanContinue = False
  bsShowMessage("Esta anotação está sendo usada em " + SQL.FieldByName("QTD").AsString + " beneficiário(s). Alteração não permitida", "E")
  Set SQL = Nothing
  Exit Sub
End If

SQL.Clear
SQL.Add("SELECT COUNT(HANDLE) QTD")
SQL.Add("FROM SAM_CONTRATO_ANOTADM")
SQL.Add("WHERE ANOTACAO = :HANOTACAO")
SQL.ParamByName("HANOTACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True
If SQL.FieldByName("QTD").AsInteger >0 Then
  CanContinue = False
  bsShowMessage("Esta anotação está sendo usada em " + SQL.FieldByName("QTD").AsString + " contrato(s). Alteração não permitida", "E")
  Set SQL = Nothing
  Exit Sub
End If

SQL.Clear
SQL.Add("SELECT COUNT(HANDLE) QTD")
SQL.Add("FROM SAM_FAMILIA_ANOTADM")
SQL.Add("WHERE ANOTACAO = :HANOTACAO")
SQL.ParamByName("HANOTACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True
If SQL.FieldByName("QTD").AsInteger >0 Then
  CanContinue = False
  bsShowMessage("Esta anotação está sendo usada em " + SQL.FieldByName("QTD").AsString + " família(s). Alteração não permitida", "E")
  Set SQL = Nothing
  Exit Sub
End If

SQL.Clear
SQL.Add("SELECT COUNT(HANDLE) QTD")
SQL.Add("FROM SAM_PRESTADOR_ANOTADM")
SQL.Add("WHERE ANOTACAO = :HANOTACAO")
SQL.ParamByName("HANOTACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True
If SQL.FieldByName("QTD").AsInteger >0 Then
  CanContinue = False
  bsShowmEssage("Esta anotação está sendo usada em " + SQL.FieldByName("QTD").AsString + " prestador(es). Alteração não permitida", "E")
  Set SQL = Nothing
  Exit Sub
End If

SQL.Clear
SQL.Add("SELECT COUNT(HANDLE) QTD")
SQL.Add("FROM SAM_PLANO_ANOTADM")
SQL.Add("WHERE ANOTACAO = :HANOTACAO")
SQL.ParamByName("HANOTACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True
If SQL.FieldByName("QTD").AsInteger >0 Then
  CanContinue = False
  bsShowMessage("Esta anotação está sendo usada em " + SQL.FieldByName("QTD").AsString + " plano(s). Alteração não permitida", "E")
  Set SQL = Nothing
  Exit Sub
End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "RELATORIO") Then
		RELATORIO_OnBtnClick
	End If
End Sub

'Public Sub TIPO_OnChange()
'	CurrentQuery.UpdateRecord
'	TABLE_AfterScroll
'End Sub
