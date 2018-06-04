'HASH: 7AD7F1B2ECE3462CA9E617C3205B3ED8


Public Sub CID_OnPopup(ShowPopup As Boolean)
  CID.AnyLevel = True
  Dim OLEAutorizador As Object
  Dim handlexx As Long
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  ShowPopup = False
  handlexx = OLEAutorizador.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CID").Value = handlexx
  End If
  Set OLEAutorizador = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim PRINCIPAL As String
  PRINCIPAL = CurrentQuery.FieldByName("CidPrincipal").AsString

  If Not InTransaction Then StartTransaction

  If PRINCIPAL = "S" Then
    Dim qUpdate As Object
    Set qUpdate = NewQuery
    qUpdate.Active = False
    qUpdate.Clear
    qUpdate.Add("Update At_Atend_Cid set CidPrincipal='N' where atendimento=:atendimento")
    qUpdate.ParamByName("atendimento").Value = CurrentQuery.FieldByName("atendimento").Value
    qUpdate.ExecSQL
    Set qUpdate = Nothing
  End If

  If InTransaction Then Commit
End Sub

