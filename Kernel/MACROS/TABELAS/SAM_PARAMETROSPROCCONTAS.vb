'HASH: 4D0BE78DD4D85622B3A27A3F835EEE66
'Macro: SAM_PARAMETROSPROCCONTAS
'#uses "*bsShowMessage"

Public Sub NUMERACAOPEG_OnChange()
  'If CurrentQuery.FieldByName("NUMERACAOPEG").AsString="M" Then
  '  CurrentQuery.FieldByName("SUGEREPEGCARTAREMESSA").Value=False
  'End If


End Sub

Public Sub TABLE_AfterScroll()

  If VisibleMode Then
    MASCARABRASINDICE.LocalWhere = "HANDLE NOT IN (@MASCARABRASINDICEGENERICO , @MASCARABRASINDICERESTHOSPMARCA , @MASCARABRASINDICERESTHOSPGEN)"
    MASCARABRASINDICEGENERICO.LocalWhere = "HANDLE NOT IN (@MASCARABRASINDICE , @MASCARABRASINDICERESTHOSPMARCA , @MASCARABRASINDICERESTHOSPGEN)"
    MASCARABRASINDICERESTHOSPMARCA.LocalWhere = "HANDLE NOT IN (@MASCARABRASINDICE , @MASCARABRASINDICEGENERICO , @MASCARABRASINDICERESTHOSPGEN)"
    MASCARABRASINDICERESTHOSPGEN.LocalWhere = "HANDLE NOT IN (@MASCARABRASINDICE , @MASCARABRASINDICEGENERICO , @MASCARABRASINDICERESTHOSPMARCA)"
  End If

  Set vDllBSPro006 = CreateBennerObject("BSPRO006.Rotinas")

  TABTIPOINTERFACEPEG.Visible = vDllBSPro006.ClientePermitirDigitacaoPeg(CurrentSystem)

  Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("UTILIZACALENDARIODIARIO").AsString = "N" Then
    If CurrentQuery.FieldByName("CALENDARIOEXCECAO").AsString <> "G" Then
      bsShowMessage("Data no calendário deve ser Geral.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If ((CurrentQuery.FieldByName("TABCONTROLEPAGAMENTO").AsInteger = 1)  And (CurrentQuery.FieldByName("HABILITAVERIFICACAOPEG").AsString <> "S")) Then
    bsShowMessage("Necessário habilitar situação de verificação do PEG para utilizar o Controle de pagamento.", "E")
    CanContinue = False
    Exit Sub
  End If
  If ((Not (CurrentQuery.FieldByName("TPSERVICOLISTAPOSITIVA").IsNull) And (CurrentQuery.FieldByName("TPSERVICOLISTANEGATIVA").IsNull Or CurrentQuery.FieldByName("TPSERVICOLISTANEUTRA").IsNull)) Or _
  	 (Not (CurrentQuery.FieldByName("TPSERVICOLISTANEGATIVA").IsNull) And (CurrentQuery.FieldByName("TPSERVICOLISTAPOSITIVA").IsNull Or CurrentQuery.FieldByName("TPSERVICOLISTANEUTRA").IsNull)) Or _
  	 (Not (CurrentQuery.FieldByName("TPSERVICOLISTANEUTRA").IsNull) And (CurrentQuery.FieldByName("TPSERVICOLISTANEGATIVA").IsNull Or CurrentQuery.FieldByName("TPSERVICOLISTAPOSITIVA").IsNull))) Then
	bsShowMessage("Para utilizar o tipo de serviço por lista positiva, negativa ou neutra necessário o preenchimento dos 3 parâmetros.", "E")
	TPSERVICOLISTAPOSITIVA.SetFocus
    CanContinue = False
    Exit Sub
  End If
  If((CurrentQuery.FieldByName("TABCOMUNICADEVOLUCAO").AsInteger = 1) And ((CurrentQuery.FieldByName("MENSAGEMDEVOLUCAO").IsNull) Or (CurrentQuery.FieldByName("MENSAGEMDEVOLUCAOGUIA").IsNull)))Then
	bsshowmessage("Para utilizar o envio de comunicado na devolução do PEG e da Guia necessário informar as mensagens utilizadas como padrão.", "E")
	MENSAGEMDEVOLUCAO.SetFocus
	CanContinue = False
    Exit Sub
  End If

End Sub


Public Sub UTILIZACALENDARIODIARIO_OnChange()

  If CurrentQuery.FieldByName("UTILIZACALENDARIODIARIO").AsString = "S" Then
    CurrentQuery.FieldByName("calendarioexcecao").AsString = "G"
  End If

End Sub

