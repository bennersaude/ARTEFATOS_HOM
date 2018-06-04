'HASH: 612A605C3F195C34BE6A33AF92562F58

'#uses "*bsShowMessage"
Option Explicit

Public Sub BOTAOGERAREVENTOS_OnClick()

  Dim Duplica As Object
  Set Duplica = CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem, "SAM_TAXAADMPAGAMENTO_EVENTO", "TAXAADMINISTRATIVA", CurrentQuery.FieldByName("HANDLE").AsInteger, "Gerando eventos")
  Set Duplica = Nothing

  RefreshNodesWithTable "SAM_TAXAADMPAGAMENTO_EVENTO"
End Sub

'Não é permitido alterar o Tipo de Taxa posteriormente
Public Sub TABLE_AfterScroll()
  AjustarVisibilidadeDoBotaoGerarEventos

  If (CurrentQuery.State = 3) Then 'Inserção

	TIPODETAXA.ReadOnly = False
	DATAINICIAL.ReadOnly = False

  Else

  	TIPODETAXA.ReadOnly = True

  	If TaxaJaUtilizada() Then
  		DATAINICIAL.ReadOnly = True
  	Else
  		DATAINICIAL.ReadOnly = False
  	End If

  End If
  
End Sub

Private Sub AjustarVisibilidadeDoBotaoGerarEventos()
  If CurrentQuery.FieldByName("TIPODETAXA").AsInteger = 1 Then 'Tipo Linear
    BOTAOGERAREVENTOS.Visible = True
  Else
    BOTAOGERAREVENTOS.Visible = False
  End If
End Sub

Private Function TaxaJaUtilizada()

  TaxaJaUtilizada = False

  If (CurrentQuery.FieldByName("JAUTILIZADA").AsBoolean = True) Then
  	TaxaJaUtilizada = True
  End If

End Function

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If TaxaJaUtilizada() Then
  	bsShowMessage("O registro não pode ser alterado pois já possui vínculo com faturamento(s).", "E")
  	CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not ValidarDataDeFechamento Then
  	bsShowMessage("A Data de fechamento não pode ser anterior a data atual.", "E")
  	CanContinue = False
  	Exit Sub
  End If

  Dim vsMensagemErro As String
  Dim callEntity As CSEntityCall

  If CurrentQuery.FieldByName("TIPODETAXA").AsInteger = 1 Then
  	'Tipo de taxa linear
    Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPrestadorTaxaAdmPagamento, Benner.Saude.Entidades", "ValidatingTaxaLinear")

    callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)
	vsMensagemErro = CStr(callEntity.Execute)

  Else
  	'Tipo de taxa Escalonada
	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPrestadorTaxaAdmPagamento, Benner.Saude.Entidades", "ValidatingTaxaEscalonada")

	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
    callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)
	vsMensagemErro = CStr(callEntity.Execute)
  End If

  Set callEntity =  Nothing

  If vsMensagemErro <> "" Then
    bsShowMessage(vsMensagemErro, "E")
    CanContinue = False
  End If

End Sub

Private Function ValidarDataDeFechamento

  ValidarDataDeFechamento = True

  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    Exit Function
  End If

  If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < Date Then
	ValidarDataDeFechamento = False
  End If

End Function

Public Sub TIPODETAXA_OnChange()
  	CurrentQuery.FieldByName("PERCENTUAL").Value = 0.00
End Sub
