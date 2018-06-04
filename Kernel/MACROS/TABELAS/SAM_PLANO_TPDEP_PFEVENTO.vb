'HASH: 07F527D868D769DAEC7F2827E5D465EF
'SAM_PLANO_TPDEP_PFEVENTO

Public Sub PLANOPFEVENTO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_PLANO_PFEVENTO|SAM_TABPF[SAM_PLANO_PFEVENTO.TABELAPFEVENTO = SAM_TABPF.HANDLE]", "DESCRICAO", 1, "Descrição", "PLANO = " + Str(RecordHandleOfTable("SAM_PLANO")), "Procura por PF", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PLANOPFEVENTO").Value = handlexx
  End If
  Set Procura = Nothing

End Sub

