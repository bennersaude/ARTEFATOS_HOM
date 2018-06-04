'HASH: 1E6F0DFD34EA08312A2CA1FC04A33469
'Macro: SAM_TGE_GRAU
'#Uses "*bsShowMessage"
Option Explicit

Public Sub BOTAOEXCLUIRCOMPL_OnClick()
  Dim SQL As Object
  Dim SQLConsulta As Object
  Set SQL = NewQuery
  Set SQLConsulta = NewQuery
  Dim handleComplementar As Long

  SQLConsulta.Clear
  SQLConsulta.Add("SELECT HANDLE FROM SAM_TGE_COMPLEMENTAR WHERE EVENTOAGERAR = :EVENTO AND GRAUAGERAR = :GRAU")
  SQLConsulta.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQLConsulta.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").AsInteger
  SQLConsulta.Active = True

  If Not SQLConsulta.EOF Then
    handleComplementar = SQLConsulta.FieldByName("HANDLE").AsInteger
  End If

  If Not InTransaction Then StartTransaction
  SQL.Add("DELETE FROM SAM_TGE_COMPLEMENTAR WHERE EVENTOAGERAR = :EVENTO AND GRAUAGERAR = :GRAU")
  SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").AsInteger
  SQL.ExecSQL
  If InTransaction Then Commit

  If (handleComplementar > 0) Then
    WriteAudit("E", HandleOfTable("SAM_TGE_COMPLEMENTAR"), handleComplementar, "Exclusão de Evento Complementar")
  End If

  If VisibleMode Then
    RefreshNodesWithTable "SAM_TGE_COMPLEMENTAR"
  End If

  Set SQL = Nothing
  Set SQLConsulta = Nothing
End Sub

Public Sub BOTAOINCLUIRCOMPL_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.RequestLive = True
  SQL.Add("SELECT * FROM SAM_TGE_COMPLEMENTAR A WHERE A.EVENTO = :EVENTO AND A.EVENTOAGERAR = :EVENTOAGERAR AND A.GRAUAGERAR = :GRAUAGERAR")
  SQL.ParamByName("EVENTO").Value = RecordHandleOfTable("SAM_TGE")
  SQL.ParamByName("EVENTOAGERAR").Value = RecordHandleOfTable("SAM_TGE")
  SQL.ParamByName("GRAUAGERAR").Value = CurrentQuery.FieldByName("GRAU").AsInteger
  SQL.Active = True
  If SQL.EOF Then
    SQL.Insert
    SQL.FieldByName("HANDLE").Value = NewHandle("SAM_TGE_COMPLEMENTAR")
    SQL.FieldByName("EVENTO").Value = RecordHandleOfTable("SAM_TGE")
    SQL.FieldByName("EVENTOAGERAR").Value = RecordHandleOfTable("SAM_TGE")
    If Not CurrentQuery.FieldByName("GRAU").IsNull Then
      SQL.FieldByName("GRAUAGERAR").Value = CurrentQuery.FieldByName("GRAU").Value
    End If
    SQL.FieldByName("NAOACEITAPFINFAUT").Value = "N"
    SQL.FieldByName("QTD").Value = 1
    SQL.Post

	WriteAudit("I", HandleOfTable("SAM_TGE_COMPLEMENTAR"), SQL.FieldByName("HANDLE").AsInteger, "Inclusão de Evento Complementar")

    If VisibleMode Then
      RefreshNodesWithTable "SAM_TGE_COMPLEMENTAR"
    End If

  End If
  Set SQL = Nothing
End Sub


' A TGE não possui mais o grau principal

Public Sub TABLE_AfterDeletexxxx()
  GravarGrauPrincipalTGE
End Sub

Public Sub TABLE_AfterDelete()
  GravarGrauPrincipalTGE
End Sub

Public Sub TABLE_AfterInsert()
  'Eduardo - 20/01/2005 - SMS 38276
  'Verifica se deixa visível ou não os campos com informações do prazo intervalar
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT TABPRAZOINTERVALAR FROM SAM_TGE WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True
  'Se prazo intervalar for por grau deixa visível campos para preenchimento
  If SQL.FieldByName("TABPRAZOINTERVALAR").AsInteger = 2 Then
    TABTIPOPERIODOINTERVALAR.Visible = True
  Else
    TABTIPOPERIODOINTERVALAR.Visible = False
  End If

  Set SQL = Nothing
  'fim SMS 38276
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Dim SQL2 As Object
  Set SQL2 = NewQuery
  Set SQL = NewQuery

  SQL.Add("SELECT COUNT(*) NREC FROM SAM_TGE_GRAU WHERE EVENTO=:EVENTO")
  SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True
  If (SQL.FieldByName("NREC").AsInteger = 1) Or (CurrentQuery.FieldByName("GRAUPRINCIPAL").AsString = "S") Then
    GRAUPRINCIPAL.ReadOnly = True
  Else
    GRAUPRINCIPAL.ReadOnly = False
  End If

  SQL2.Active = False
  SQL2.Clear
  SQL2.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :PHANDLE")
  SQL2.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL2.Active = True

  CODIGO.Text = ""
  CODIGO.Text = "Estrutura: " + SQL2.FieldByName("ESTRUTURA").AsString

  'Fábio 29/01/03
  SQL.Clear
  SQL.Add("SELECT TABPRAZOINTERVALAR FROM SAM_TGE WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True
  If SQL.FieldByName("TABPRAZOINTERVALAR").AsInteger = 2 Then
    TABTIPOPERIODOINTERVALAR.Visible = True
  Else
    TABTIPOPERIODOINTERVALAR.Visible = False
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim vNovoGrauPrincipal, vHandleGrau As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT COUNT (HANDLE) TOT FROM SAM_TGE_COMPLEMENTAR A WHERE A.EVENTOAGERAR = :EVENTO AND A.GRAUAGERAR = :GRAU AND EVENTO = :EVENTO")
  SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").AsInteger
  SQL.Active = True
  If SQL.FieldByName("TOT").AsInteger>0 Then
    CanContinue = False
    bsShowMessage("O registro não pode ser excluído por estar cadastrado como evento complementar", "E")
  End If

  If CurrentQuery.FieldByName("GRAUPRINCIPAL").AsString = "S" Then
    SQL.Clear
    SQL.Add("SELECT HANDLE, EVENTO, GRAU FROM SAM_TGE_GRAU WHERE EVENTO = :EVENTO AND HANDLE <> :EVENTOAEXCLUIR")
    SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
    SQL.ParamByName("EVENTOAEXCLUIR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      vNovoGrauPrincipal = SQL.FieldByName("HANDLE").AsInteger
      vHandleGrau = SQL.FieldByName("GRAU").AsInteger
      SQL.Clear
      SQL.Add("UPDATE SAM_TGE_GRAU SET GRAUPRINCIPAL = 'S' WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = vNovoGrauPrincipal
      SQL.ExecSQL

      SQL.Clear
      SQL.Add("UPDATE SAM_TGE SET GRAUPRINCIPAL = :pGRAU WHERE HANDLE = :pHANDLE")
      SQL.ParamByName("pGRAU").AsInteger = vHandleGrau
      SQL.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
      SQL.ExecSQL
    Else
      SQL.Clear
      SQL.Add("UPDATE SAM_TGE")
      SQL.Add("   SET GRAUPRINCIPAL = NULL")
      SQL.Add(" WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_TGE")
      SQL.ExecSQL
    End If
  End If

  SQL.Clear
  SQL.Add("SELECT SUM(QTD) QTDPACOTES FROM (")
  SQL.Add("SELECT COUNT(HANDLE) QTD")
  SQL.Add("  FROM SAM_PCTNEGESTADO")
  SQL.Add(" WHERE EVENTO = :EVENTO")
  SQL.Add("UNION")
  SQL.Add("SELECT COUNT(HANDLE) QTD")
  SQL.Add("  FROM SAM_PCTNEGFILIAL")
  SQL.Add(" WHERE EVENTO = :EVENTO")
  SQL.Add("UNION")
  SQL.Add("SELECT COUNT(HANDLE) QTD")
  SQL.Add("  FROM SAM_PCTNEGGERAL")
  SQL.Add(" WHERE EVENTO = :EVENTO")
  SQL.Add("UNION")
  SQL.Add("SELECT COUNT(HANDLE) QTD")
  SQL.Add("  FROM SAM_PCTNEGMUNIC")
  SQL.Add(" WHERE EVENTO = :EVENTO")
  SQL.Add("UNION")
  SQL.Add("SELECT COUNT(HANDLE) QTD")
  SQL.Add("  FROM SAM_PCTNEGPREST")
  SQL.Add(" WHERE EVENTO = :EVENTO")
  SQL.Add("UNION")
  SQL.Add("SELECT COUNT(HANDLE) QTD")
  SQL.Add("  FROM SAM_PCTNEGREDE")
  SQL.Add(" WHERE EVENTO = :EVENTO")
  SQL.Add(") X")

  SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").Value
  SQL.Active = True

  Dim vQtd As Integer
  vQtd = SQL.FieldByName("QTDPACOTES").AsInteger

  If vQtd > 0 Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT ORIGEMVALOR FROM SAM_GRAU WHERE HANDLE = :GRAU")
    SQL.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").Value
    SQL.Active = True

    If SQL.FieldByName("ORIGEMVALOR").AsInteger = 7 Then 'Pacote
      CanContinue = False
      bsShowMessage("O registro não pode ser excluído por haver " + Str(vQtd) + " pacotes negociados para este evento.", "E")
    End If

  End If

  Set SQL = Nothing
End Sub


' A TGE não possui mais o grau principal

Public Sub TABLE_AfterPost
  Dim SQL As Object
  Set SQL = NewQuery
  If CurrentQuery.FieldByName("GRAUPRINCIPAL").AsString = "S" Then
    SQL.Clear
    SQL.Add("UPDATE SAM_TGE_GRAU SET GRAUPRINCIPAL = 'N' WHERE EVENTO = :EVENTO AND HANDLE <> :HANDLE")
    SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").Value
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.ExecSQL
  End If
  GravarGrauPrincipalTGE
  Set SQL = Nothing
End Sub

Public Sub MarcarGrauPrincipalTGE(HandleGrau As Long)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("UPDATE SAM_TGE SET GRAUPRINCIPAL = :GRAUPRINCIPAL WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.ParamByName("GRAUPRINCIPAL").Value = HandleGrau
  SQL.ExecSQL
  Set SQL = Nothing
End Sub

Public Sub DesmarcarGrauPrincipalTGE
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("UPDATE SAM_TGE SET GRAUPRINCIPAL = NULL WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.ExecSQL
  Set SQL = Nothing
End Sub

Public Sub GravarGrauPrincipalTGE
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = :EVENTO AND GRAUPRINCIPAL = 'S'")
  SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    MarcarGrauPrincipalTGE(SQL.FieldByName("GRAU").AsInteger) 'Alterar o campo GRAU da SAM_TGE
  Else
    DesmarcarGrauPrincipalTGE
  End If
  Set SQL = Nothing
End Sub


'#Uses "*ProcuraGrau"

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  If ShowPopup = False Then
    Exit Sub
  End If
  '  If Len(GRAU.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraGrau(GRAU.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  '  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE FROM SAM_TGE_GRAU WHERE EVENTO=:EVENTO AND GRAU=:GRAU AND HANDLE<>:HANDLE")
  SQL.ParamByName("GRAU").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
  SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    bsShowMessage("Já foi incluído este grau como válido para o evento", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If



  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(*) NREC FROM SAM_TGE_GRAU WHERE EVENTO=:EVENTO")
  SQL.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True
  If SQL.FieldByName("NREC").AsInteger = 0 Then
    CurrentQuery.FieldByName("GRAUPRINCIPAL").AsString = "S"
  End If

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT PRONTUARIO FROM SAM_TGE WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True
  If (SQL.FieldByName("PRONTUARIO").AsString<>"T") And (CurrentQuery.FieldByName("IGNORAPRONTUARIO").AsString = "S") Then
    bsShowMessage("Para marcar ignorar prontuário, deve o evento estar marcado que todos os graus irão para prontuário", "E")
    CanContinue = False
    CurrentQuery.FieldByName("IGNORAPRONTUARIO").AsString = "N"
  End If

  Set SQL = Nothing
End Sub
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOEXCLUIRCOMPL"
			BOTAOEXCLUIRCOMPL_OnClick
		Case "BOTAOINCLUIRCOMPL"
			BOTAOINCLUIRCOMPL_OnClick
	End Select
End Sub
