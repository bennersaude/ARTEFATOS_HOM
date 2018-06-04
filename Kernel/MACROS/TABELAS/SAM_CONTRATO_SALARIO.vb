'HASH: E42D772FAB42886B52E75DD75F6CA007
'Macro da tabela: SAM_CONTRATO_SALARIO


Public Sub TABLE_AfterPost()
  Dim upd As Object
  Set upd = NewQuery

  upd.Add("UPDATE SAM_BENEFICIARIO_SALARIO     ")
  upd.Add("   SET COMPLEMENTAR = :COMPLEMENTAR,")
  upd.Add("       SEQUENCIA    = :SEQUENCIA    ")
  upd.Add(" WHERE CONTRATOSALARIO = :HANDLE    ")
  upd.ParamByName("HANDLE"      ).AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  upd.ParamByName("COMPLEMENTAR").AsString  = CurrentQuery.FieldByName("COMPLEMENTAR").AsString
  upd.ParamByName("SEQUENCIA"   ).AsString  = CurrentQuery.FieldByName("SEQUENCIA").AsString
  upd.ExecSQL

  Set upd = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
    Dim BSBen003Dll As Object

    Set BSBen003Dll = CreateBennerObject("BSBen003.Salario")
    CurrentQuery.FieldByName("SEQUENCIA").AsString = BSBen003Dll.GeraSeqSalarioContrato(CurrentSystem, _
    								   												 	CurrentQuery.FieldByName("CONTRATO").AsInteger, _
    								   													CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
    Set BSBen003Dll = Nothing
End Sub
