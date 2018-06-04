'HASH: C91541B632C8D233C46A852F2DC7EA86
'Macro da tabela: SAM_CONTRATO_AUTOGESTAO
'#Uses "*bsShowMessage"

Option Explicit

Dim aux As String

Public Sub BOTAORECALCULAR_OnClick()
    If CurrentQuery.FieldByName("DATARECALCULO").IsNull Then

		If bsShowMessage( "Tabela de Contribuições de Faixas Etárias Finalizadas?", "Q") = vbYes Then

			Dim SamRotinaRecalcMensal As Object

			If VisibleMode Then
				Set SamRotinaRecalcMensal = CreateBennerObject("BSINTERFACE0065.RotinaRecalculo")
				SamRotinaRecalcMensal.RecalcContratoAutoGestao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
			Else
				Dim vsMensagemErro As String
				Dim viRetorno As Long
				Dim viHandleRotina As Integer

				Set SamRotinaRecalcMensal = CreateBennerObject("SamRecalcMensal.Rotinas")

				viHandleRotina = SamRotinaRecalcMensal.CriarRotinaRecalculoAutoGestao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 0)

				If viHandleRotina > 0 Then
		    		Set SamRotinaRecalcMensal = Nothing
					Set SamRotinaRecalcMensal = CreateBennerObject("BSServerExec.ProcessosServidor")
					viRetorno = SamRotinaRecalcMensal.ExecucaoImediata(CurrentSystem, _
										 "SamRecalcMensal", _
										 "Rotinas", _
										 "Processamento (Recálculo Auto-Gestão) - Rotina: " + CStr(viHandleRotina), _
										 viHandleRotina, _
										 "SAM_ROTINARECALCULOMENSALID", _
										 "SITUACAOPROCESSAMENTO", _
										 "", _
										 "", _
										 "P", _
										 False, _
										 vsMensagemErro, _
										 Null)

					If viRetorno = 0 Then
						bsShowMessage("A rotina foi enviada para execução no servidor", "I")
					Else
						bsShowMessage("Erro ao enviar rotina para o servidor: " + vsMensagemErro, "I")
					End If

					Set SamRotinaRecalcMensal = Nothing

		    		WriteAudit("P", HandleOfTable("SAM_ROTINARECALCULOMENSALID"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Recálculo de Auto-Gestão - Processamento da Rotina")

					RefreshNodesWithTable("SAM_ROTINARECALCULOMENSALID")

				Else
					bsShowMessage("Erro ao criar rotina.", "I")
				End If

			End If

		End If
	Else
		bsShowMessage("O recálculo já foi processado!", "I")
	End If
End Sub

Public Sub PESSOASUPLEMENTACAOPF_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle   As Long
  Dim vCampos   As String
  Dim vColunas  As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "NOME"

  vCriterio = "TABFISICAJURIDICA = 2 AND EHCONVENIO = 'S' "

  vCampos = "Nome"

  vHandle = interface.Exec(CurrentSystem, "SFN_PESSOA", vColunas, 1, vCampos, vCriterio, "Pessoa para Suplementação de PF", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PESSOASUPLEMENTACAOPF").Value = vHandle
  End If

  Set interface = Nothing

End Sub


Public Sub TABLE_AfterScroll()
  aux = CurrentQuery.FieldByName("permitesuplementacaopf").AsString
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vsMensagem As String

  Dim qModulo As Object
  Set qModulo = NewQuery

  qModulo.Clear
  qModulo.Add("SELECT MOD.TIPOMODULO             ")
  qModulo.Add("  FROM SAM_MODULO       MOD,      ")
  qModulo.Add("       SAM_CONTRATO_MOD COM,      ")
  qModulo.Add("       SAM_CONTRATO     CON       ")
  qModulo.Add(" WHERE MOD.HANDLE   = COM.MODULO  ")
  qModulo.Add("   AND CON.HANDLE   = COM.CONTRATO")
  qModulo.Add("   AND COM.CONTRATO = :CONTRATO   ")
  qModulo.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qModulo.Active = True

  If (qModulo.FieldByName("TIPOMODULO").AsString = "S") Then
    If ((aux = "S") And (aux <> CurrentQuery.FieldByName("PERMITESUPLEMENTACAOPF").AsString)) Then
      vsMensagem = "Não é permitido desmarcar se já existirem, para o contrato, módulos de suplementação de PF."
      bsShowMessage(vsMensagem, "E")
      CanContinue = False
      Set qModulo = Nothing
      Exit Sub
    End If
  End If
  Set qModulo = Nothing

  Dim qCompetencia As Object
  Set qCompetencia = NewQuery

  qCompetencia.Clear
  qCompetencia.Add("SELECT HANDLE                    ")
  qCompetencia.Add("  FROM SAM_CONTRATO_AUTOGESTAO   ")
  qCompetencia.Add(" WHERE CONTRATO    = :CONTRATO   ")
  qCompetencia.Add("   AND COMPETENCIA = :COMPETENCIA")
  qCompetencia.Add("   AND HANDLE     <> :HANDLE     ")
  qCompetencia.ParamByName("CONTRATO"   ).AsInteger  = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qCompetencia.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  qCompetencia.ParamByName("HANDLE"     ).AsInteger  = CurrentQuery.FieldByName("HANDLE").AsInteger
  qCompetencia.Active = True

  If (Not qCompetencia.EOF) Then
    vsMensagem = "Já existe um registro nesta competência."
    bsShowMessage(vsMensagem, "E")
    CanContinue = False
    Set qCompetencia = Nothing
    Exit Sub
  End If
  Set qCompetencia = Nothing

  'Permitir salvar o registro somente se, pelo menos, um dos parâmetros de
  'base de cota patronal e contribuição social dos titulares estiver marcado.
  'De acordo com a SMS 60048.
  If ((CurrentQuery.FieldByName("SALARIOEHBASECPCS").AsString = "N") And _
      (CurrentQuery.FieldByName("SALARIONORMALEHBASECPCS").AsString = "N") And _
      (CurrentQuery.FieldByName("APOSENTADORIAEHBASECPCS").AsString = "N") And _
      (CurrentQuery.FieldByName("APOSENTADORIACOMPEHBASECPCS").AsString = "N") And _
      (CurrentQuery.FieldByName("OUTRASRENDASEHBASECPCS").AsString = "N")) Then
    bsShowMessage("Pelo menos um dos parâmetros de base de cota patronal e contribuição social deve estar marcado.", "E")
    CanContinue = False
    Exit Sub
  End If

  'No caso de estar marcado o campo FATURARCPCSDECIMOTERCEIRO ou o campo
  'FATURARCPCSDECIMOQUARTO, será obrigatório informar um valor para o campo COMPETENCIAFATCPCSDTDQ.
  'Também deve verificar se tem, pelo menos, um dos parâmetros de base de cota patronal e contribuição social marcado.
  'De acordo com a SMS 60048.
  If ((CurrentQuery.FieldByName("FATURARCPCSDECIMOTERCEIRO").AsString = "S") Or (CurrentQuery.FieldByName("FATURARCPCSDECIMOQUARTO").AsString = "S")) Then
    If (CurrentQuery.FieldByName("COMPETENCIAFATCPCSDTDQ").IsNull) Then
      vsMensagem = ""
      vsMensagem = vsMensagem + "O contrato está configurado para faturar cota patronal e contribuição social sobre 13º e 14º salários."
      vsMensagem = vsMensagem + Chr(13) + Chr(10)
      vsMensagem = vsMensagem + "É obrigatório informar a competência para faturamento."

      bsShoWMessage(vsMensagem, "E")
      CanContinue = False
      Exit Sub
    End If

    If ((CurrentQuery.FieldByName("SALARIOEHBASECPCSDTDQ").AsString = "N") And _
        (CurrentQuery.FieldByName("SALARIONORMALEHBASECPCSDTDQ").AsString = "N") And _
        (CurrentQuery.FieldByName("APOSENTADORIAEHBASECPCSDTDQ").AsString = "N") And _
        (CurrentQuery.FieldByName("APOSENTADORIACEHBASECPCSDTDQ").AsString = "N") And _
        (CurrentQuery.FieldByName("OUTRASRENDASEHBASECPCSDTDQ").AsString = "N")) Then
      bsShowMessage("Pelo menos um dos parâmetros de base de cota patronal e contribuição social para 13º e 14º salários deve estar marcado.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If ((CurrentQuery.FieldByName("PRIMEIRAPARCELA").AsString = "3") And (CurrentQuery.FieldByName("SEGUNDAPARCELA").AsString = "2")) Then
    bsShowMessage("As duas primeiras parcelas não podem ser proporcionais simultaneamente.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAORECALCULAR" Then
		BOTAORECALCULAR_OnClick
	End If
End Sub
