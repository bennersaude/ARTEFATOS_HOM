'HASH: AE084208DCACBC2DFAD4A872D90D7FBE
'#Uses "*VerificarBloqueioAlteracoes"
'#Uses "*VerificarBloqueioAlteracoesReapresentacao"
'#Uses "*bsShowMessage"
'#Uses "*RecordHandleOfTableInterfacePEG"

Public Sub TABLE_AfterScroll()
	Dim qConsulta As Object
	Set qConsulta   = NewQuery
	qConsulta.Clear

    qConsulta.Add("SELECT DATACONTABIL FROM SAM_PEG WHERE HANDLE = :HANDLE ")
    qConsulta.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    qConsulta.Active = True

	If(CurrentQuery.State =2)Or(CurrentQuery.State =3)Then
		CurrentQuery.FieldByName("NOVADATACONTABIL").AsDateTime = qConsulta.FieldByName("DATACONTABIL").AsDateTime
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
  AlteracaoPeg.AddParameter(pdtDateTime, CurrentQuery.FieldByName("NOVADATACONTABIL").AsDateTime)
  AlteracaoPeg.AddParameter(pdtInteger, CurrentQuery.FieldByName("MOTIVO").AsInteger)
  AlteracaoPeg.AddParameter(pdtString, CurrentQuery.FieldByName("OBSERVACOES").AsString)
  AlteracaoPeg.Execute("AlterarDataContabil")

  Set AlteracaoPeg = Nothing
End Sub
