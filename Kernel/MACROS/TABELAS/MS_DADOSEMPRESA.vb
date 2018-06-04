'HASH: FF7D6D7C861A53F68FE327368A67302E

Public Sub BENEFICIARIO_OnChange()
  Dim qMatricula As Object
  Set qMatricula = NewQuery

  qMatricula.Clear
  qMatricula.Add("SELECT MATRICULA FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO")
  qMatricula.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qMatricula.Active = True

  CurrentQuery.FieldByName("MATRICULA").AsInteger = qMatricula.FieldByName("MATRICULA").AsInteger

  qMatricula.Active = False
End Sub

