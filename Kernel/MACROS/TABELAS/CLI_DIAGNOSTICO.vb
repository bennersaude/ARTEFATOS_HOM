'HASH: 115498C9F4BD016100A0A248E2F5F71C
'CLI_DIAGNOSTICO

Public Sub CID_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("CLI_DIAGNOSTICO")

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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim PRINCIPAL As String
  PRINCIPAL = CurrentQuery.FieldByName("CIDPRINCIPAL").AsString
  If PRINCIPAL = "S" Then
    Dim qUpdate As Object
    Set qUpdate = NewQuery
    qUpdate.Active = False
    qUpdate.Clear
    qUpdate.Add("UPDATE CLI_DIAGNOSTICO SET CIDPRINCIPAL = 'N'")
    qUpdate.Add(" WHERE ATENDIMENTO = :ATENDIMENTO")
    qUpdate.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("ATENDIMENTO").Value
    qUpdate.ExecSQL
    Set qUpdate = Nothing
  End If
End Sub

