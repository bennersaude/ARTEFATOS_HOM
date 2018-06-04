'HASH: 5CCD7743448B805374276498E0BDA719
'Macro: SAM_BENEFICIARIO_CARTAOIDENTIF
'#Uses "*bsShowMessage"
'#Uses "*RegistrarLogAlteracao"

' Mauricio Ibelli -04/05/2001 -sms 2236 -Nao permitir exclusao em cartao que ja possue Autorizacao
Option Explicit
Dim vMascaraBeneficiario As String
Dim continuar As Boolean

Public Function CartaoEmAutoriz()

  Dim SQL As Object

  CartaoEmAutoriz = True

  Set SQL = NewQuery

  SQL.Add("SELECT COUNT(*) CONTADOR")
  SQL.Add("  FROM SAM_AUTORIZ ")
  SQL.Add(" WHERE BENEFICIARIO = :BENEFICIARIO")
  SQL.Add("   AND DV = :DV")
  SQL.Add("   AND SITUACAO <> 'C'")
  SQL.ParamByName("BENEFICIARIO").Value = RecordHandleOfTable("SAM_BENEFICIARIO")
  SQL.ParamByName("DV").Value = CurrentQuery.FieldByName("DV").AsInteger

  SQL.Active = True

  If SQL.FieldByName("contador").Value <>0 Then
    bsShowMessage("O cartão possue Autorização. Exclusão não permitida.", "E")
    Exit Function
  End If

  CartaoEmAutoriz = False

End Function

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("A Tabela não pode estar em edição", "E")
    Exit Sub
  End If

If ((VisibleMode) Or (WebMode And WebMenuCode = "")) Then
  If CurrentQuery.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Cartão já cancelado.", "E")
    continuar = False
    Exit Sub
  Else
  	continuar = True
  End If
End If



If (VisibleMode) Or (WebMode And WebMenuCode <> "")Then
  Dim Interface As Object
  If bsShowMessage("Confirma o cancelamento do cartão ?", "Q") = vbYes Then

       Dim vsMensagemErro As String
       Dim viRetorno As Integer
       Dim vvContainer As CSDContainer

	   Set vvContainer = NewContainer
       SessionVar("HANDLE") = CurrentQuery.FieldByName("HANDLE").AsString

	   Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")

	   If VisibleMode Then

       		viRetorno = Interface.Exec(CurrentSystem, _
        									1, _
            	                            "TV_FORM0001", _
                	                        "Cancelamento de Cartão", _
                    	                    0, _
                        	                250, _
                            	            300, _
                                	        False, _
                                    	    vsMensagemErro, _
                                        	vvContainer)

	     	Select Case viRetorno
		      	Case -1
	      			bsShowMessage("Operação cancelada pelo usuário!", "I")
     	  		Case  0
	     	  		'bsShowMessage("Opção selecionada" + vvContainer.Field("OPCAO").AsString , "I")
	     	  		RegistrarLogAlteracao "SAM_BENEFICIARIO_CARTAOIDENTIF", CurrentQuery.FieldByName("HANDLE").AsInteger, "BOTAOCANCELAR_OnClick"
     	  		Case  1
	     	  		bsShowMessage(vsMensagemErro, "I")
    		End Select
    	End If
     Set Interface = Nothing

    RefreshNodesWithTable("SAM_BENEFICIARIO_CARTAOIDENTIF")
  End If
End If
End Sub

Public Sub BOTAODESBLOQUEAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("A Tabela não pode estar em edição", "I")
    Exit Sub
  End If
  If(CurrentQuery.FieldByName("SITUACAO").Value <>"B")Then
  	bsShowMessage("Situação não permite Desbloqueio.", "I")
  	Exit Sub
  End If
If(CurrentQuery.FieldByName("DATAGERACAO").IsNull)Then
	bsShowMessage("Cartão não gerado.", "I")
	Exit Sub
End If

Dim Interface As Object

If bsShowMessage("Confirma o desbloqueio do cartão ?", "Q") = vbYes Then
  Set Interface = CreateBennerObject("BSBEN009.Beneficiario")
  bsShowMessage(Interface.DesbloqueiaCartaoIndividual(CurrentSystem), "I")
  Set Interface = Nothing
  RefreshNodesWithTable("SAM_BENEFICIARIO_CARTAOIDENTIF")
End If

End Sub
Public Sub BOTAODESVINCULAR_OnClick()
  Dim vlog As String
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT R.USUARIOEXPORTACAO")
  SQL.Add("  FROM SAM_ROTINACARTAO_CARTAO RC, ")
  SQL.Add("       SAM_ROTINACARTAO R")
  SQL.Add(" WHERE RC.CARTAOIDENTIFICACAO = :HCARTAOIDENTIFICACAO")
  SQL.Add("   AND R.HANDLE = RC.ROTINACARTAO")
  SQL.ParamByName("HCARTAOIDENTIFICACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.FieldByName("USUARIOEXPORTACAO").IsNull Then
    bsShowMessage("Os cartões já foram exportados. Não é prossível desvincular o cartão", "I")
    Exit Sub
  End If

  If bsShowMessage("Confirma a desvinculação do cartão ?", "Q") = vbYes Then


    If Not InTransaction Then StartTransaction

  	SQL.Clear
  	SQL.Add("UPDATE SAM_ROTINACARTAO_CARTAO SET")
  	SQL.Add("    ROTINACARTAO = NULL")
  	SQL.Add("WHERE CARTAOIDENTIFICACAO = :HCARTAOIDENTIFICACAO")
  	SQL.ParamByName("HCARTAOIDENTIFICACAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  	SQL.ExecSQL

  	If InTransaction Then Commit

	  Set SQL = Nothing

  	RefreshNodesWithTable("SAM_BENEFICIARIO_CARTAOIDENTIF")

  	End If

End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
  Dim qDadosFiltro As Object
  Set qDadosFiltro = NewQuery
  qDadosFiltro.Clear
  qDadosFiltro.Add("SELECT CONTRATO, FAMILIA, HANDLE FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
  qDadosFiltro.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qDadosFiltro.Active =True

  UserVar("FORM.21.NONE.2") = "CONTRATO=" + qDadosFiltro.FieldByName("CONTRATO").AsString +"|FAMILIA="+ qDadosFiltro.FieldByName("FAMILIA").AsString + "|BENEFICIARIO="+ qDadosFiltro.FieldByName("HANDLE").AsString + "|"

  Set qDadosFiltro = Nothing

  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("A Tabela não pode estar em edição","I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAEMISSAO").IsNull Then
    'MsgBox("O Cartão já foi impresso")
    'Exit Sub
  End If

  'SMS24760 -Relatório Cartões Provisórios
  Dim QueryBuscaHandle As Object
  Dim qBuscaRelatorio As Object
  Dim qVerificaContrato As Object
  Dim Especifico As Object

  Set QueryBuscaHandle = NewQuery
  Set qBuscaRelatorio = NewQuery
  Set qVerificaContrato = NewQuery
  Set Especifico = CreateBennerObject("ESPECIFICO.uEspecifico")

  qVerificaContrato.Active = False
  qVerificaContrato.Add("SELECT C.DEFINIRTIPOCARTAONAGERACAO, B.EHTITULAR ")
  qVerificaContrato.Add("  FROM SAM_CONTRATO C,              ")
  qVerificaContrato.Add("       SAM_BENEFICIARIO B           ")
  qVerificaContrato.Add(" WHERE C.HANDLE = B.CONTRATO        ")
  qVerificaContrato.Add("   AND B.HANDLE = :BENEF            ")
  qVerificaContrato.ParamByName("BENEF").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qVerificaContrato.Active = True

  UserVar("HANDLEBENEF") = CStr(CurrentQuery.FieldByName("BENEFICIARIO").AsInteger)

  Dim samBeneficiarioBLL As CSBusinessComponent
  Dim retorno As String
  Dim RelatorioHandle As Long

  Set samBeneficiarioBLL = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.SamBeneficiarioCartaoIdentifBLL, Benner.Saude.Beneficiarios.Business")
  samBeneficiarioBLL.AddParameter(pdtInteger, CLng(CurrentQuery.FieldByName("HANDLE").AsString))
  retorno = samBeneficiarioBLL.Execute("ValidaImpressaoProvisoriaCartao")
  If retorno <> "" Then
    On Error GoTo retornoMensagem
      RelatorioHandle = CInt(retorno)
      SessionVar("HCARTAOIDENTIF") = CurrentQuery.FieldByName("HANDLE").AsString
	  ReportPreview(RelatorioHandle, "", False, False)
	  GoTo FinalizarImpressao

    retornoMensagem:
      bsShowMessage(retorno, "I")
	FinalizarImpressao:
      Finalizar(qBuscaRelatorio, QueryBuscaHandle, qVerificaContrato)
	  Exit Sub
  End If
  'sms 22944 - fernando
  If qVerificaContrato.FieldByName("DEFINIRTIPOCARTAONAGERACAO").AsString = "1" Then

    QueryBuscaHandle.Active = False
    QueryBuscaHandle.Add("SELECT T.RELATORIOESPECIFICO, T.IMPRIMIRDEPENDENTES  ")
    QueryBuscaHandle.Add("  FROM SAM_TIPOCARTAO T      ")
    QueryBuscaHandle.Add("  JOIN SAM_CONTRATO_TIPOCARTAO CT ON CT.TIPOCARTAO = T.HANDLE      ")
    QueryBuscaHandle.Add(" WHERE CT.HANDLE = :TIPOCARTAO ")
    QueryBuscaHandle.ParamByName("TIPOCARTAO").AsInteger = CurrentQuery.FieldByName("TIPOCARTAO").AsInteger
    QueryBuscaHandle.Active = True

    If (Especifico.BEN_GeraCartao(CurrentSystem)) And _
       (QueryBuscaHandle.FieldByName("IMPRIMIRDEPENDENTES").AsString = "S") And (qVerificaContrato.FieldByName("EHTITULAR").AsString = "N") Then
        bsShowMessage("Não pode ser impresso por causa do tipo de cartão, parâmetro marcado para imprimir cartão somente para seu titular!", "I")
        Set Especifico = Nothing
        Finalizar(qBuscaRelatorio, QueryBuscaHandle, qVerificaContrato)
		Exit Sub
    End If

    Set Especifico = Nothing

    If QueryBuscaHandle.FieldByName("RELATORIOESPECIFICO").IsNull Then
      bsShowMessage("Falta configurar o relatório de impressão no tipo de cartão do beneficiário!", "I")
      Finalizar(qBuscaRelatorio, QueryBuscaHandle, qVerificaContrato)
	  Exit Sub
    End If

    RelatorioHandle = QueryBuscaHandle.FieldByName("RELATORIOESPECIFICO").AsInteger

    SessionVar("HCARTAOIDENTIF") = "0"
  End If


  ReportPreview(RelatorioHandle, "", False, False)

  Set samBeneficiarioBLL = Nothing

  Finalizar(qBuscaRelatorio, QueryBuscaHandle, qVerificaContrato)

End Sub

Private Function Finalizar(qBuscaRelatorio As Object, QueryBuscaHandle As Object, qVerificaContrato As Object)
	Set qBuscaRelatorio = Nothing
    Set QueryBuscaHandle = Nothing
    Set qVerificaContrato = Nothing
End Function

Public Sub BOTAOREGISTRORETIRADA_OnClick()

   If VisibleMode Then
       Dim vsMensagemErro As String
       Dim viRetorno As Integer
       Dim vvContainer As CSDContainer
       Dim Interface As Object

	   Set vvContainer = NewContainer
       SessionVar("HANDLE") = CurrentQuery.FieldByName("HANDLE").AsString

	   Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")

	   If VisibleMode Then

       		viRetorno = Interface.Exec(CurrentSystem, _
        									1, _
            	                            "TV_FORM0008", _
                	                        "Registrar retirada do Cartão", _
                    	                    0, _
                        	                200, _
                            	            610, _
                                	        False, _
                                    	    vsMensagemErro, _
                                        	vvContainer)

	     	Select Case viRetorno
		      	Case -1
	      			bsShowMessage("Operação cancelada pelo usuário!", "I")
     	  		Case  0
	     	  		'bsShowMessage("Opção selecionada" + vvContainer.Field("OPCAO").AsString , "I")
     	  		Case  1
	     	  		bsShowMessage(vsMensagemErro, "I")
    		End Select
    	End If
     Set Interface = Nothing
End If

End Sub
Public Sub BOTAOREGULARIZAR_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("A Tabela não pode estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "B" Then
    bsShowMessage("Situação não permite regularização.", "I")
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery


  Dim Interface As Object
  Set Interface = CreateBennerObject("BSBEN009.Beneficiario")
  bsShowMessage(Interface.Regularizar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger), "I")

  RefreshNodesWithTable("SAM_BENEFICIARIO_CARTAOIDENTIF")
End Sub
Public Sub TABLE_AfterCancel()

End Sub

Public Sub TABLE_AfterScroll()

'If(Not CurrentQuery.FieldByName("DATAEMISSAO").IsNull)Or(CurrentQuery.FieldByName("SITUACAO").AsString <>"B")Then
 COMPETENCIAFINALVALIDADE.ReadOnly = True
 COMPETENCIAINICIALVALIDADE.ReadOnly = True
 DATAINICIALVALIDADE.ReadOnly = True
 DATAFINALVALIDADE.ReadOnly = True
'Else
'COMPETENCIAFINALVALIDADE.ReadOnly = False
'COMPETENCIAINICIALVALIDADE.ReadOnly = False
'DATAINICIALVALIDADE.ReadOnly = False
'DATAFINALVALIDADE.ReadOnly = False
'End If

Dim query1 As Object

Set query1 = NewQuery

query1.Clear
query1.Add("SELECT B.CONTRATANTE ")
query1.Add("  FROM SAM_BENEFICIARIO A,")
query1.Add("       SAM_CONTRATO B WHERE A.HANDLE = :HBENEFICIARIO")
query1.Add("  AND B.HANDLE = A.CONTRATO")
query1.ParamByName("HBENEFICIARIO").Value = RecordHandleOfTable("SAM_BENEFICIARIO")'CurrentQuery.FieldByName("handle").AsInteger

query1.Active = True

LBCONTRATANTE.Text = "Contratante: " + query1.FieldByName("CONTRATANTE").AsString

'Ferreira - SMS 43213 - 16.08.2005
query1.Active = False
query1.Clear
query1.Add("SELECT VIGENCIACARTAO FROM SAM_PARAMETROSBENEFICIARIO ")
query1.Active = True
If query1.FieldByName("VIGENCIACARTAO").AsString = "C" Then  ' Alteração para WEB SMS - 95224
  GRUPOVALIDADE.Visible = True
  GRUPOVALIDADEDATA.Visible = False
  COMPETENCIAFINALVALIDADE.Required = True
  DATAFINALVALIDADE.Required = False

Else
  GRUPOVALIDADE.Visible = False
  GRUPOVALIDADEDATA.Visible = True
  COMPETENCIAFINALVALIDADE.Required = False
  DATAFINALVALIDADE.Required = True
End If
Set query1 = Nothing
'Final SMS 43213

UserVar("HANDLEBENEF") = CStr(CurrentQuery.FieldByName("BENEFICIARIO").AsInteger)

End Sub
Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  If CurrentQuery.FieldByName("VIA").AsInteger >1 Then
    CanContinue = False
    bsShowMessage("Apenas a primeira via pode ser excluída", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <>"B" Then
    CanContinue = False
    bsShowMessage("Apenas os cartões bloqueados podem ser excluídos", "E")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAEMISSAO").IsNull Then
    CanContinue = False
    bsShowMessage("O cartão já foi emitido. Exclusão não permitida", "E")
    Exit Sub
  End If

  ' sms 2236
  If CartaoEmAutoriz Then
    CanContinue = False
    Exit Sub
  End If

  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT HANDLE")
  SQL.Add("FROM SAM_BENEFICIARIO_CARTAOIDENTIF")
  SQL.Add("WHERE BENEFICIARIO = :HBENEFICIARIO")
  SQL.Add("  AND HANDLE <> :HCARTAOATUAL")
  SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.ParamByName("HCARTAOATUAL").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Existem outros cartões para o beneficiário. Exclusão não permitida", "E")
    Set SQL = Nothing
    Exit Sub
  End If

  SQL.Clear
  SQL.Add("DELETE FROM SAM_ROTINACARTAO_CARTAO")
  SQL.Add("WHERE CARTAOIDENTIFICACAO = :HCARTAO")
  SQL.ParamByName("HCARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL

  RegistrarLogAlteracao "SAM_BENEFICIARIO", CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "TABLE_BeforeDelete"

  Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID

		Case "BOTAODESBLOQUEAR"
			BOTAODESBLOQUEAR_OnClick

		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick

		Case "BOTAODESVINCULAR"
			BOTAODESVINCULAR_OnClick

		Case "BOTAOIMPRIMIR"
			BOTAOIMPRIMIR_OnClick

		Case "BOTAOREGISTRORETIRADA"
			BOTAOREGISTRORETIRADA_OnClick

		Case "BOTAOREGULARIZAR"
			BOTAOREGULARIZAR_OnClick

		Case "CANCELAR"
			CanContinue = CancelarCartao
	End Select
End Sub

Function CancelarCartao() As Boolean
	BOTAOCANCELAR_OnClick
	CancelarCartao = continuar
End Function
