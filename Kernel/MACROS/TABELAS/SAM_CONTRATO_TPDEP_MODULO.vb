'HASH: AA258B0CA629E9690CE87E4CC9BE603E
'#Uses "*bsShowMessage"
'SAM_CONTRATO_TPDEP_MODULO

Public Sub CONTRATOMODULO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long
  Dim sql As Object


  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_CONTRATO_MOD.MODULO = SAM_MODULO.HANDLE]", "DESCRICAO", 1, "Descrição", "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO")), "Procura por Módulo", True, "")
  If handlexx <= 0 Then
    Exit Sub
  Else

    Set sql = NewQuery
    sql.Active = False
    sql.Clear
    sql.Add("SELECT CONTRATOMODULO FROM SAM_CONTRATO_TPDEP_MODULO WHERE CONTRATOMODULO = :CONTRATOMODULO AND CONTRATOTPDEP = :CONTRATOTPDEP")
    sql.ParamByName("CONTRATOMODULO").AsInteger = handlexx
    sql.ParamByName("CONTRATOTPDEP").AsInteger = CurrentQuery.FieldByName("CONTRATOTPDEP").AsInteger
    sql.Active = True

    If Not sql.FieldByName("CONTRATOMODULO").IsNull Then
      bsShowMessage("Este módulo já está cadastrado.", "I")
      Exit Sub
    End If


    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOMODULO").Value = handlexx
  End If
  Set Procura = Nothing

End Sub

