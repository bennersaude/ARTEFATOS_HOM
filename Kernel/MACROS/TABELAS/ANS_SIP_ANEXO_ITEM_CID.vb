'HASH: B5D000E09A7910BF049FFD6830E02BF2
'Macro: ANS_SIP_ANEXO_ITEM_CID
'#Uses "*bsShowMessage"

Public Sub BOTAOEXCLUIRCID_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.ExcluirCID(CurrentSystem, "ANS_SIP_ANEXO_ITEM_CID", "Excluindo CID", "SAM_CID", "CID", "SIPANEXO", CurrentQuery.FieldByName("SIPANEXO").AsInteger, "S", "")
  Set Obj = Nothing
End Sub

Public Sub BOTAOINCLUIRCID_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.GerarCID(CurrentSystem, "ANS_SIP_ANEXO_ITEM_CID", "Gerando CID", "SAM_CID", "CID", "SIPANEXO", CurrentQuery.FieldByName("SIPANEXO").AsInteger, "S", "")
  Set Obj = Nothing
End Sub

Public Sub CID_OnPopup(ShowPopup As Boolean)
  Dim interface As Object ' SMS 78297 - Willian - 23/03/2007
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|DESCRICAO"
  vCriterio = ""
  vCampos = "Estrutura|Descrição"
  vHandle = interface.Exec(CurrentSystem, "SAM_CID", vColunas, 1, vCampos, vCriterio, "CID", True, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CID").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CanContinue Then
    If CurrentQuery.FieldByName("IDADEINICIAL").IsNull Then
      If Not CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
        bsShowMessage("Informe a Idade Inicial ou apague a Idade Final.", "E")
        CanContinue = False
      End If
    Else
      If CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
        bsShowMessage("Informe a Idade Final ou apague a Idade Inicial.", "E")
        CanContinue = False
      Else
        If CurrentQuery.FieldByName("IDADEINICIAL").AsInteger > CurrentQuery.FieldByName("IDADEFINAL").AsInteger Then
          bsShowMessage("A Idade Final deve ser maior que a Idade Inicial", "E")
          CanContinue = False
        End If
      End If
    End If
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOEXCLUIRCID") Then
		BOTAOEXCLUIRCID_OnClick
	End If
	If (CommandID = "BOTAOINCLUIRCID") Then
		BOTAOINCLUIRCID_OnClick
	End If
End Sub
