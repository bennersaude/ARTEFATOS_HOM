'HASH: C5C02B3B2C34886DFA955171B39A8C29
'macro SAM_REDERESTRITA
'#Uses "*bsShowMessage"


Dim vgCodigo As Long

Public Sub BOTAODUPLICAR_OnClick()
  Dim DuplicaRedeRestritaDLL As Object
  Dim pSMensagem As String

  Set DuplicaRedeRestritaDLL = CreateBennerObject("SamDupRedeRestrita.SamDupRedeRestrita")
  DuplicaRedeRestritaDLL.Executar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, pSMensagem)

  If Len(pSMensagem) > 0 Then
    bsShowMessage(pSMensagem, "I")
  End If

  Set DuplicaRedeRestritaDLL = Nothing
  RefreshNodesWithTable("SAM_REDERESTRITA")

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "E", "P", Msg)
  If vFitro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  '---claudemir ---
  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_REDERESTRITA_PRESTADOR WHERE REDERESTRITA = :REDERESTRITA")
  SQL.ParamByName("REDERESTRITA").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.Active = True

  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Operação Cancelada !!!" + Chr(10) + "Motivo: Esta rede esta cadastrada nas rede restritas do prestador.", "E")
    Exit Sub
  End If
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT * FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :HANDLE OR REDERESTRITACONTIDA = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.Active = True

  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Operação Cancelada !!!" + Chr(10) + "Motivo: Esta rede pertence a um configuração de redes contidas", "E")
    Exit Sub
  End If
  '-----------------
  Set SQL = Nothing
End Sub


Public Sub TABLE_AfterPost()
  Dim SQL As Object
  Dim UPD As Object
  Set SQL = NewQuery


  SQL.Add("SELECT COUNT(HANDLE) NREC FROM SAM_REDERESTRITA ")
  SQL.Active = True

  If(vgCodigo <>CurrentQuery.FieldByName("CODIGO").Value)And(SQL.FieldByName("NREC").AsInteger >1)Then
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_REDERESTRITA WHERE CODIGO >= :CODIGO AND HANDLE <> :HANDLE")
  SQL.ParamByName("CODIGO").Value = CurrentQuery.FieldByName("CODIGO").AsInteger
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    Set UPD = NewQuery
    UPD.Add("UPDATE SAM_REDERESTRITA ")
    UPD.Add("   SET CODIGO = CODIGO + 1         ")
    UPD.Add("WHERE HANDLE IN (SELECT HANDLE FROM SAM_REDERESTRITA WHERE CODIGO BETWEEN :CODIGO1 AND :CODIGO2 AND HANDLE <> :HANDLE)")
    UPD.ParamByName("CODIGO1").Value = CurrentQuery.FieldByName("CODIGO").AsInteger
    UPD.ParamByName("CODIGO2").Value = vgCodigo
    UPD.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    UPD.ExecSQL
    RefreshNodesWithTable("SAM_REDERESTRITA")
  End If

  Set UPD = Nothing

End If
Set SQL = Nothing
End Sub


Public Sub TABLE_AfterEdit()
  vgCodigo = CurrentQuery.FieldByName("CODIGO").Value
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(HANDLE) QTDE ")
  SQL.Add("  FROM SAM_REDERESTRITA   ")
  SQL.Add(" WHERE HANDLE <> :PHANDLE ")
  SQL.Add("   AND CODIGO = :PCODIGO  ")
  SQL.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("PCODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Já existe uma rede restrita com o código informado!", "E")
    CurrentQuery.FieldByName("CODIGO").Clear
    CODIGO.SetFocus
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAODUPLICAR") Then
		BOTAODUPLICAR_OnClick
	End If
End Sub
