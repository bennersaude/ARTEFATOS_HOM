'HASH: 9C7CF7DF75860FE9A80ED5E4F18AB0DD
'#Uses "*bsShowMessage"
'#Uses "*VerificarBloqueioAlteracoes"
'#Uses "*VerificarBloqueioAlteracoesReapresentacao"
'#Uses "*RecordHandleOfTableInterfacePEG"

Public Sub TABLE_AfterScroll()
	Dim qConsulta As Object
	Set qConsulta   = NewQuery
	qConsulta.Clear
	qConsulta.Active = False

    qConsulta.Add("SELECT TOTALINFORMADO, TOTALPAGARINFORMADO FROM SAM_PEG WHERE HANDLE = :HANDLE ")
    qConsulta.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    qConsulta.Active = True

    If(CurrentQuery.State =2)Or(CurrentQuery.State =3)Then
		CurrentQuery.FieldByName("NOVOVALORAPRESENTADO").AsFloat = qConsulta.FieldByName("TOTALINFORMADO").AsFloat
		CurrentQuery.FieldByName("NOVOVALORAPAGARINFORMADO").AsFloat = qConsulta.FieldByName("TOTALPAGARINFORMADO").AsFloat
	End If

    Set qConsulta   = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  If VerificarBloqueioAlteracoesReapresentacao(RecordHandleOfTableInterfacePEG("SAM_PEG")) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG é de reapresentação. ", "E")
    CanContinue = False
	Exit Sub
  End If

If VerificarBloqueioAlteracoes(RecordHandleOfTableInterfacePEG("SAM_PEG")) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG está vinculado a um agrupador de pagamento com documentos fiscais conciliados. ", "E")
    CanContinue = False
	Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim AlteracaoPeg As CSBusinessComponent

  Set AlteracaoPeg = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.SamPegBLL, Benner.Saude.ProcessamentoContas.Business")
  AlteracaoPeg.AddParameter(pdtInteger, RecordHandleOfTableInterfacePEG("SAM_PEG"))
  AlteracaoPeg.AddParameter(pdtDecimal, CurrentQuery.FieldByName("NOVOVALORAPRESENTADO").AsFloat)
  AlteracaoPeg.AddParameter(pdtDecimal, CurrentQuery.FieldByName("NOVOVALORAPAGARINFORMADO").AsFloat)
  AlteracaoPeg.AddParameter(pdtInteger, CurrentQuery.FieldByName("MOTIVO").AsInteger)
  AlteracaoPeg.AddParameter(pdtString, CurrentQuery.FieldByName("OBSERVACOES").AsString)
  AlteracaoPeg.Execute("AlterarValorApresentado")

  Set AlteracaoPeg = Nothing
End Sub
