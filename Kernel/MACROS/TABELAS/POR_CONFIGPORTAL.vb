'HASH: 584A0DB427AFB581A5164CDF7F68038D
'MACRO: POR_CONFIGPORTAL
'#Uses "*bsShowMessage"

Option Explicit

Dim exibeTipoPrestadorFiltroConsulta As Boolean
Dim redeAtedimentoPropriaAtual As Boolean
Dim redeAtedimentoPorPlano As Boolean


Public Sub BOTAOPROCLICENCA_OnClick()
  On Error GoTo Erro
	Dim serverExec As CSServerExec

	Set serverExec = NewServerExec

	serverExec.Description = "Processar Licença Portal "
	serverExec.DllClassName = "Benner.Saude.Web.PortalServicos.ManagerInstaladorFuncionalidades.Processar"

	serverExec.Execute

	bsShowMessage("Processamento da licença enviado para o servidor. Aguarde alguns minutos antes de acessar o Portal Serviços novamente.", "I")
	GoTo Finaliza
  Erro:
    bsShowMessage(Err.Description, "E")

  Finaliza:

End Sub


Public Sub CONSIDERADADOSBANCOBENEF_OnChange()
  If Not (CurrentQuery.FieldByName("CONSIDERADADOSBANCOBENEF").AsBoolean) Then
     CONSIDERADADOSDEPENDENTE.Visible = False
  Else
	 CONSIDERADADOSDEPENDENTE.Visible = True
  End If
End Sub

Public Sub EXIBTIPOPRESTFILTROCONS_OnChange()
  AtualizarCamposTipoPrestador(Not exibeTipoPrestadorFiltroConsulta)
End Sub

Public Sub PERMITEENVIONFPROTOC_OnChange()
	CurrentQuery.FieldByName("TIPODOCPADRAOENVIONF").AsString = ""
	CurrentQuery.FieldByName("PERMITEENVIONFPEGSAGRUPADOS").AsString = "N"

End Sub

Public Sub REDEATENDPORPLANO_OnChange()
  AtualizarRedeAtendimentoPorPlano(Not redeAtedimentoPorPlano)
End Sub

Public Sub REDEATENDPROPRIA_OnChange()
  AtualizarRedeAtendimentoPropria(Not redeAtedimentoPropriaAtual)
End Sub

Public Sub TABLE_AfterScroll()
  Dim qParametro As Object
  Set qParametro = NewQuery

  qParametro.Active = False
  qParametro.Clear
  qParametro.Add("SELECT HANDLE FROM POR_CONFIGPORTAL ")
  qParametro.Active = True

  BOTAOPROCLICENCA.Enabled = Not qParametro.EOF

  If Not (CurrentQuery.FieldByName("CONSIDERADADOSBANCOBENEF").AsBoolean) Then
     CONSIDERADADOSDEPENDENTE.Visible = False
  End If

  Set qParametro = Nothing

  AtualizarRedeAtendimentoPropria(CurrentQuery.FieldByName("REDEATENDPROPRIA").AsBoolean)
  AtualizarCamposTipoPrestador(CurrentQuery.FieldByName("EXIBTIPOPRESTFILTROCONS").AsBoolean)
  AtualizarRedeAtendimentoPorPlano(CurrentQuery.FieldByName("REDEATENDPORPLANO").AsBoolean)


End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  AtualizarRedeAtendimentoPropria(CurrentQuery.FieldByName("REDEATENDPROPRIA").AsBoolean)
  AtualizarCamposTipoPrestador(CurrentQuery.FieldByName("EXIBTIPOPRESTFILTROCONS").AsBoolean)
  AtualizarRedeAtendimentoPorPlano(CurrentQuery.FieldByName("REDEATENDPORPLANO").AsBoolean)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  ENDEROREDEATEND.ReadOnly = False
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("BLOQUEIOTENTATIVAACESSO").AsString  = "1" Then
		If CurrentQuery.FieldByName("QUANTIDADETENTATIVAS").IsNull Then
			bsShowMessage("Campo quantidade é obrigatório","E")
			CanContinue = False
		End If
	End If

	If CurrentQuery.FieldByName("PERMITEENVIONFPROTOC").AsString = "1" Then
    	If CurrentQuery.FieldByName("PERMITEENVIONFPEGSAGRUPADOS").IsNull Then
			bsShowMessage("Campo permite envio de NF para PEGs agrupados é obrigatório","E")
			CanContinue = False
        End If

    	If CurrentQuery.FieldByName("TIPODOCPADRAOENVIONF").IsNull Then
			bsShowMessage("Campo tipo de documento padrão para envio de NF é obrigatório","E")
			CanContinue = False
        End If
    End If

End Sub

Public Sub AtualizarRedeAtendimentoPropria(checked As Boolean)
  redeAtedimentoPropriaAtual = checked
  ENDEROREDEATEND.ReadOnly = Not redeAtedimentoPropriaAtual
End Sub

Public Sub AtualizarCamposTipoPrestador(checked As Boolean)
  exibeTipoPrestadorFiltroConsulta = checked
  OBRIGTIPOPRESTADOR.ReadOnly = Not exibeTipoPrestadorFiltroConsulta
End Sub

Public Sub AtualizarRedeAtendimentoPorPlano(checked As Boolean)
  redeAtedimentoPorPlano = checked
  GRPREDEATENDIMENTOPLANO.Visible = redeAtedimentoPorPlano
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "CMDPROCESSARLICENCA"
		BOTAOPROCLICENCA_OnClick
	End Select
End Sub
