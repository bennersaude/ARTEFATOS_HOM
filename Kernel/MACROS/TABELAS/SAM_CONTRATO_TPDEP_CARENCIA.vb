'HASH: C3433E359EE45360763341D90CF059BD
'SAM_CONTRATO_TPDEP_CARENCIA
'#Uses "*bsShowMessage"

Public Sub CONTRATOCARENCIA_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long
  Dim sql As Object


  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_CARENCIA|SAM_CARENCIA[SAM_CONTRATO_CARENCIA.CARENCIA = SAM_CARENCIA.HANDLE]", "DESCRICAO", 1, "Descrição", "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO")), "Procura por Carência", True, "")
  If handlexx <= 0 Then
    Exit Sub
  Else

    Set sql = NewQuery

    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOCARENCIA").Value = handlexx

    sql.Active = False
    sql.Clear
    sql.Add("SELECT CONTRATOCARENCIA FROM SAM_CONTRATO_TPDEP_CARENCIA WHERE CONTRATOCARENCIA = :CONTRATOCARENCIA")
    sql.ParamByName("CONTRATOCARENCIA").Value = handlexx
    sql.Active = True
    If Not sql.FieldByName("CONTRATOCARENCIA").IsNull Then
      bsShowMessage("Esta carência já está cadastrado.", "I")
      Exit Sub
    End If

  End If
  Set Procura = Nothing

End Sub

