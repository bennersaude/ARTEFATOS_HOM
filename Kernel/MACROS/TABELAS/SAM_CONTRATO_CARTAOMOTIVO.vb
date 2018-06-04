'HASH: 11243A1A13ADD9D6AE3D285617D20597
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object

  CanContinue = False

  Set SQL = NewQuery

  SQL.Add("SELECT handle FROM SAM_ROTINACARTAO_CONTRATO")
  SQL.Add("WHERE CARTAOMOTIVOEMISSAO=:MOTIVO and contrato = :CONTRATO")

  SQL.ParamByName("CONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
  SQL.ParamByName("MOTIVO").Value = CurrentQuery.FieldByName("CARTAOMOTIVO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Contrato/Motivo usado em parâmetros para geração de Cartão. Exclusão não permitida.", "I")
    SQL.Active = False
    Exit Sub
  End If

  Set SQL = Nothing

  CanContinue = True

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("COBRAREMISSAO").Value = "S" Then
    If CurrentQuery.FieldByName("TAXACARTAO").Value = 0 Then
      bsShowMessage("Os cartões desse contrato para esse motivo, não serão faturados!", "I")
    End If
  End If

  If CurrentQuery.FieldByName("COBRAREMISSAO").Value = "N" Then
    If CurrentQuery.FieldByName("TAXACARTAO").Value >0 Then
      bsShowMessage("O valor para faturamento do cartão a ser utilizado, será dos parâmetros gerais!", "I")
    End If
  End If

  Dim qSel As Object
  Set qSel = NewQuery

  qSel.Active = False
  qSel.Clear
  qSel.Add("SELECT HANDLE ")
  qSel.Add("  FROM SAM_CONTRATO_CARTAOMOTIVO")
  qSel.Add(" WHERE CONTRATO = :CONTRATO AND CARTAOMOTIVO = :CARTAOMOTIVO AND HANDLE <> :HANDLE")
  qSel.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qSel.ParamByName("CARTAOMOTIVO").AsInteger = CurrentQuery.FieldByName("CARTAOMOTIVO").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.Active = True

  If Not qSel.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Motivo de emissão do cartão já cadastrado para o contrato.", "E")
    CanContinue = False
    Set qSel = Nothing
    Exit Sub
  End If

  Set qSel = Nothing

End Sub

