'HASH: 775AF9911265966A9CC9DA143C2818F8

Public Sub PLANOMODULO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_PLANO_MOD|SAM_MODULO[SAM_PLANO_MOD.MODULO = SAM_MODULO.HANDLE]", "DESCRICAO", 1, "Descrição", "PLANO = " + Str(RecordHandleOfTable("SAM_PLANO")), "Procura por Módulo", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PLANOMODULO").Value = handlexx
  End If
  Set Procura = Nothing

End Sub

