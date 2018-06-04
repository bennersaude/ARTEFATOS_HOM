'HASH: F391D9EBC2B2317D9E4C7CEAD75711DF
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qAux As Object
  Set qAux = NewQuery

  qAux.Clear
  qAux.Add("SELECT COUNT(1) QTD                                 ")
  qAux.Add("  FROM SAM_BENEFICIARIO_PFEVENTO_ESPE               ")
  qAux.Add(" WHERE HANDLE <> :HANDLE                            ")
  qAux.Add(" AND BENEFICIARIOPFEVENTO = :BENEFICIARIOPFEVENTO   ")
  qAux.Add(" AND ESPECIALIDADE = :ESPECIALIDADE                 ")

  qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qAux.ParamByName("BENEFICIARIOPFEVENTO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIOPFEVENTO").AsInteger
  qAux.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  qAux.Active = True

  If qAux.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Especialidade já cadastrada !", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

