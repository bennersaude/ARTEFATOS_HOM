'HASH: 6E809C608567D46A9A4FE50D43904A32
'MACRO.GER_ITEM_CID


Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.GerarCID(CurrentSystem, "GER_ITEM_CID", "Gerando CID", "SAM_CID", "CID", "ITEM", CurrentQuery.FieldByName("ITEM").AsInteger, "S", "")
  Set Obj = Nothing
End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.ExcluirCID(CurrentSystem, "GER_ITEM_CID", "Excluindo CID", "SAM_CID", "CID", "ITEM", CurrentQuery.FieldByName("ITEM").AsInteger, "S", "")
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

Public Sub TABLE_AfterScroll()
  Dim qSQL As Object
  Set qSQL = NewQuery
  qSQL.Add("SELECT INTERNACAO")
  qSQL.Add("  FROM GER_ITEM")
  qSQL.Add(" WHERE HANDLE = :ITEM")
  qSQL.ParamByName("ITEM").AsInteger = RecordHandleOfTable("GER_ITEM")
  qSQL.Active = True
  MOTIVOALTA.Visible = (qSQL.FieldByName("INTERNACAO").AsString = "S")
  Set qSQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CanContinue Then
    If CurrentQuery.FieldByName("IDADEINICIAL").IsNull Then
      If Not CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
        CanContinue = False
        MsgBox("Informe a Idade Inicial ou apague a Idade Final.")
      End If
    Else
      If CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
        CanContinue = False
        MsgBox("Informe a Idade Final ou apague a Idade Inicial.")
      Else
        If CurrentQuery.FieldByName("IDADEINICIAL").AsInteger > CurrentQuery.FieldByName("IDADEFINAL").AsInteger Then
          CanContinue = False
          MsgBox("A Idade Final deve ser maior que a Idade Inicial")
        End If
      End If
    End If
  End If
End Sub




