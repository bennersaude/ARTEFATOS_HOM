'HASH: D36501986F5545DB236FC2766D37F58A
'SAM_CONTRATO_TPDEP_PFEVENTO
'#Uses "*bsShowMessage"

Public Sub CONTRATOPFEVENTO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long
  Dim SQL As Object


  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_PFEVENTO|SAM_TABPF[SAM_CONTRATO_PFEVENTO.TABELAPFEVENTO = SAM_TABPF.HANDLE]", "DESCRICAO", 1, "Descrição", "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO")), "Procura por PF", True, "")
  If handlexx <= 0 Then
    Exit Sub
  Else
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOPFEVENTO").Value = handlexx


    Set SQL = NewQuery
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT CONTRATOPFEVENTO FROM SAM_CONTRATO_TPDEP_PFEVENTO WHERE CONTRATOPFEVENTO = :CONTRATOPFEVENTO")
    SQL.ParamByName("CONTRATOPFEVENTO").Value = handlexx
    SQL.Active = True

    If Not SQL.FieldByName("CONTRATOPFEVENTO").IsNull Then
      bsShowMessage("Este evento já está cadastrado.", "I")
      Exit Sub
    End If
  End If
  Set Procura = Nothing

End Sub

