'HASH: 6AFDB35FA68A43199034A894B4F634E0
'#Uses "*bsShowMessage"
'#Uses "*VerificarBloqueioAlteracoes"
'#Uses "*RecordHandleOfTableInterfacePEG"

Public Sub TABLE_AfterScroll()
	Dim qConsulta As Object
	Set qConsulta   = NewQuery
	qConsulta.Clear

    qConsulta.Add("SELECT DATAPAGAMENTO FROM SAM_PEG WHERE HANDLE = :HANDLE ")
    qConsulta.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    qConsulta.Active = True
	If(CurrentQuery.State =2)Or(CurrentQuery.State =3)Then
		CurrentQuery.FieldByName("NOVADATAPAGAMENTO").AsDateTime = qConsulta.FieldByName("DATAPAGAMENTO").AsDateTime
	End If

    Set qConsulta   = Nothing
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If VerificarBloqueioAlteracoes(RecordHandleOfTableInterfacePEG("SAM_PEG")) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG está vinculado a um agrupador de pagamento com documentos fiscais conciliados. ", "E")
    CanContinue = False
	Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim agrupadorFechado      As Boolean
  Dim qPeg                  As Object
  Dim qPagamento            As Object
  Dim qParametrosProcContas As Object
  Dim vDataNaoPermitida     As Boolean

  agrupadorFechado = VerificaAgrupadorPagamentoFechado

  If (agrupadorFechado) Then
    BsShowMessage("Não é permitida a alteração de data de pagamento do PEG que está ligado à registro de pagamento fechado.","E")
    CanContinue = False
    Exit Sub
  End If

  vDataNaoPermitida = False
  If Not CurrentQuery.FieldByName("NOVADATAPAGAMENTO").IsNull Then
	Set qPeg                  = NewQuery
	Set qPagamento            = NewQuery
	Set qParametrosProcContas = NewQuery

    qParametrosProcContas.Clear
    qParametrosProcContas.Add("SELECT UTILIZACALENDARIODIARIO FROM SAM_PARAMETROSPROCCONTAS")
    qParametrosProcContas.Active = True

	qPeg.Clear
    qPeg.Add("SELECT TABREGIMEPGTO FROM SAM_PEG WHERE HANDLE = :HANDLE ")
    qPeg.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
    qPeg.Active = True

    qPagamento.Clear
    If qPeg.FieldByName("TABREGIMEPGTO").AsInteger = 1 Then
      qPagamento.Add("SELECT DATAPROCESSAMENTO, DATAFECHAMENTO")
      qPagamento.Add("  FROM SAM_PAGAMENTO")
      qPagamento.Add(" WHERE DATAPAGAMENTO = :DATAPGTO")
      qPagamento.Add(" ORDER BY DATAFECHAMENTO")
    Else
      qPagamento.Add("SELECT DATAPROCESSAMENTO, DATAFECHAMENTO")
      qPagamento.Add("  FROM SAM_CALENDARIOREEMBOLSO")
      qPagamento.Add(" WHERE DATAPAGAMENTO = :DATAPGTO")
    End If
    qPagamento.ParamByName("DATAPGTO").Value = CurrentQuery.FieldByName("NOVADATAPAGAMENTO").AsDateTime
    qPagamento.Active = True

    If qPagamento.EOF Then
      If qParametrosProcContas.FieldByName("UTILIZACALENDARIODIARIO").AsString = "N" Then
        bsShowMessage("Data de pagamento não permitida - data não cadastrada no calendário geral.", "E")
        vDataNaoPermitida = True
      End If
    Else
      If Not qPagamento.FieldByName("DATAFECHAMENTO").IsNull Then
        bsShowMessage("Data de Pagamento não Permitida - calendário fechado.", "E")
        vDataNaoPermitida = True
      End If
    End If

    qPeg.Active                  = False
    qPagamento.Active            = False
    qParametrosProcContas.Active = False
    Set qPeg                   = Nothing
    Set qPagamento             = Nothing
    Set qPParametrosProcContas = Nothing

    Dim Interface As Object
    Set Interface = CreateBennerObject("SAMCALENDARIOPGTO.ROTINAS")
    Interface.INICIALIZAR(CurrentSystem)
    If CurrentQuery.FieldByName("NOVADATAPAGAMENTO").AsDateTime <> Interface.DIAUTILANTERIOR(CurrentSystem, CurrentQuery.FieldByName("NOVADATAPAGAMENTO").AsDateTime) Then
      bsShowMessage("Entre com um dia útil para a Data de Pagamento", "E")
      vDataNaoPermitida = True
      NOVADATAPAGAMENTO.SetFocus
    End If
    Interface.FINALIZAR
    Set Interface = Nothing

    If vDataNaoPermitida Then
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim AlteracaoPeg As CSBusinessComponent

  Set AlteracaoPeg = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.SamPegBLL, Benner.Saude.ProcessamentoContas.Business")
  AlteracaoPeg.AddParameter(pdtInteger, RecordHandleOfTableInterfacePEG("SAM_PEG"))
  AlteracaoPeg.AddParameter(pdtDateTime, CurrentQuery.FieldByName("NOVADATAPAGAMENTO").AsDateTime)
  AlteracaoPeg.AddParameter(pdtInteger, CurrentQuery.FieldByName("MOTIVO").AsInteger)
  AlteracaoPeg.AddParameter(pdtString, CurrentQuery.FieldByName("OBSERVACOES").AsString)
  AlteracaoPeg.Execute("AlterarDataPagamento")

  Set AlteracaoPeg = Nothing
End Sub

Public Function VerificaAgrupadorPagamentoFechado As Boolean
	Dim callEntity As CSEntityCall
  	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamPeg, Benner.Saude.Entidades", "VerificaPegVinculadoPagamentoFechado")
  	callEntity.AddParameter(pdtAutomatic, RecordHandleOfTableInterfacePEG("SAM_PEG"))
  	VerificaAgrupadorPagamentoFechado = CBool(callEntity.Execute)
	Set callEntity =  Nothing
End Function
