'HASH: 6C820422CA091970BBC11307BE6A0709
'#Uses "*bsShowMessage"

Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "SAM_LIMITACAO_CID", "Excluindo CID's da Limitação", "SAM_CID", "CID", "LIMITACAO", CurrentQuery.FieldByName("LIMITACAO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing
  RefreshNodesWithTable("SAM_LIMITACAO_CID")
End Sub

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "SAM_LIMITACAO_CID", "Duplicando CID's para Limitação", "SAM_CID", "CID", "LIMITACAO", CurrentQuery.FieldByName("LIMITACAO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing
  RefreshNodesWithTable("SAM_LIMITACAO_CID")
End Sub


Public Sub CID_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|DESCRICAO"
  'vCriterio =" ULTIMONIVEL = 'S'"
  vCampos = "Estrutura|Descricao"
  vHandle = interface.Exec(CurrentSystem, "SAM_CID", vColunas, 2, vCampos, vCriterio, "CID", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CID").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	 CID.AnyLevel = True
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	 CID.AnyLevel = True
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_LIMITACAO_CID                ")
Consulta.Add(" WHERE LIMITACAO = :LIMITACAO           ")
Consulta.Add("   AND CID = :CID                       ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("LIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
Consulta.ParamByName("CID").AsInteger = CurrentQuery.FieldByName("CID").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("CID já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOEXCLUIR"
			BOTAOEXCLUIR_OnClick
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
	End Select
End Sub
