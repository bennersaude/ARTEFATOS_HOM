'HASH: 7C0130F7C6C681705FEB685B14A1038E


Public Sub CID_OnPopup(ShowPopup As Boolean)
  CID.AnyLevel = True
  Dim OLEAutorizador As Object
  Dim handlexx As Long
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  ShowPopup = False
  handlexx = OLEAutorizador.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CID").Value = handlexx
  End If
  Set OLEAutorizador = Nothing
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim ATENDE As Object
  Set ATENDE = NewQuery
  ATENDE.Add("SELECT MATRICULA FROM CLI_ATENDIMENTO WHERE HANDLE = :ATENDIMENTO")
  ATENDE.ParamByName("ATENDIMENTO").Value = RecordHandleOfTable("CLI_ATENDIMENTO")
  ATENDE.Active = True
  CurrentQuery.FieldByName("MATRICULA").AsInteger = ATENDE.FieldByName("MATRICULA").AsInteger
  Set ATENDE = Nothing
End Sub

