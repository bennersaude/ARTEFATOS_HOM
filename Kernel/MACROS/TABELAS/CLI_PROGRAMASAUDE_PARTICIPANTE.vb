'HASH: 633D73BA8218B128D44B91C6560C4528
Public Sub TABLE_NewRecord()
  Dim qSelecionaMatriculaBeneficiario As Object
  Set qSelecionaMatriculaBeneficiario = NewQuery

  qSelecionaMatriculaBeneficiario.Active = False

  qSelecionaMatriculaBeneficiario.Clear
  qSelecionaMatriculaBeneficiario.Add("SELECT MATRICULA		 ")
  qSelecionaMatriculaBeneficiario.Add("  FROM SAM_BENEFICIARIO")
  qSelecionaMatriculaBeneficiario.Add(" WHERE HANDLE = :HANDLE")
  qSelecionaMatriculaBeneficiario.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO")

  qSelecionaMatriculaBeneficiario.Active = True
  CurrentQuery.FieldByName("PARTICIPANTE").AsInteger = qSelecionaMatriculaBeneficiario.FieldByName("MATRICULA").AsInteger

  Set qSelecionaMatriculaBeneficiario = Nothing
End Sub
