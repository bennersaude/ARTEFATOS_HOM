'HASH: 76066E3C6D2CF9BFBD27B9E853E2391B

Public Sub PLANOCARENCIA_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_PLANO_CARENCIA|SAM_CARENCIA[SAM_PLANO_CARENCIA.CARENCIA = SAM_CARENCIA.HANDLE]", "DESCRICAO", 1, "Descrição", "PLANO = " + Str(RecordHandleOfTable("SAM_PLANO")), "Procura por Carência", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PLANOCARENCIA").Value = handlexx
  End If
  Set Procura = Nothing

End Sub

