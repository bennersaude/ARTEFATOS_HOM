'HASH: 3609310CB5CE725A6066CBEF616CDC1B
'#Uses "*bsShowMessage"
Option Explicit
Dim vCodigoCampoComErro As String
Dim vCampoFoiALterado As Boolean
Dim vFoiALteradoDataRealizacao As Boolean


Public Sub TABLE_AfterEdit()
  If SessionVar("CODIGOCAMPOCOMERRO") <> "" Then
    vCodigoCampoComErro = SessionVar("CODIGOCAMPOCOMERRO")
    LiberarCamposParaAjuste
  End If
End Sub

Public Sub TABLE_AfterPost()
  Dim component As CSBusinessComponent
  If vFoiALteradoDataRealizacao Then
    Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Guia.Correcao, Benner.Saude.ANS.Processos")

    component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    component.Execute("AlterarDataRelizacao")

    RefreshNodesWithTable("ANS_TISMONITORAMENTO_GUIA")
  End If

  If vCampoFoiALterado Then
    Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Guia.Correcao, Benner.Saude.ANS.Processos")

    component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    component.AddParameter(pdtString, vCodigoCampoComErro)
    component.Execute("RemoverErro")

    RefreshNodesWithTable("ANS_TISMONITORAMENTO_GUIA")
  End If

  Set component = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  BloequearTodosCampos
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If SessionVar("CODIGOCAMPOCOMERRO") <> "" Then
    Dim retorno As String

    retorno = VerificarSeCampoFoiALterado

    If retorno <> "" Then
      bsShowMessage(retorno,"E")
      CanContinue = False
      Exit Sub
    End If

    If vFoiALteradoDataRealizacao Then
      If bsShowMessage("Ao alterar a data de realização da guia será removido todos os erros da guia e de seus eventos. Deseja Continuar?","Q") <> vbYes Then
        CanContinue = False
      Else
        SessionVar("CODIGOCAMPOCOMERRO") = ""
      End If
    Else
      If Not vCampoFoiALterado Then
        bsShowMessage("Não foi realizado alteração no campo com erro","E")
        CanContinue = False
      Else
        SessionVar("CODIGOCAMPOCOMERRO") = ""
    End If
    End If
  End If
End Sub

Public Sub BloequearTodosCampos
  CONTRATADOEXECUTANTECNES.ReadOnly=True
  CONTRATADOEXECUTANTECPFCNPJ.ReadOnly=True
  CONTRATADOEXECUTANTEMUNICIPIO.ReadOnly=True
  VERSAOTISSPRESTADOR.ReadOnly=True
  INDICADORENVIOPAPEL.ReadOnly=True
  BENEFSEXO.ReadOnly=True
  BENEFDATANASCIMENTO.ReadOnly=True
  BENEFMUNICIPIORESIDENCIA.ReadOnly=True
  BENEFNUMEROCNS.ReadOnly=True
  BENEFNUMEROREGISTROPLANO.ReadOnly=True
  NUMEROGUIAPRESTADOR.ReadOnly=True
  NUMEROGUIASOLICINTERNACAO.ReadOnly=True
  DATASOLICITACAO.ReadOnly=True
  DATAAUTORIZACAO.ReadOnly=True
  DATAINICIALFATURAMENTO.ReadOnly=True
  DATAFINALFATURAMENTO.ReadOnly=True
  DATAPROTOCOLO.ReadOnly=True
  DATAPAGAMENTO.ReadOnly=True
  TIPOCONSULTA.ReadOnly=True
  CBOSEXECUTANTE.ReadOnly=True
  INDICACAORECEMNATO.ReadOnly=True
  INDICACAOACIDENTE.ReadOnly=True
  CARATERATENDIMENTO.ReadOnly=True
  TIPOINTERNACAO.ReadOnly=True
  TIPOATENDIMENTO.ReadOnly=True
  MOTIVOSAIDA.ReadOnly=True
  REGIMEINTERNACAO.ReadOnly=True
  TIPOFATURAMENTO.ReadOnly=True
  CIDPRINCIPAL.ReadOnly=True
  CID2.ReadOnly=True
  CID3.ReadOnly=True
  CID4.ReadOnly=True
  NUMDIARIASACOMPANHANTE.ReadOnly=True
  NUMDIARIASUTI.ReadOnly=True
  VALORTOTALINFORMADO.ReadOnly=True
  VALORTOTALPAGOPROC.ReadOnly=True
  VALORTOTALDIARIAS.ReadOnly=True
  VALORTOTALTAXAS.ReadOnly=True
  VALORTOTALMATERIAIS.ReadOnly=True
  VALORTOTALOPME.ReadOnly=True
  VALORTOTALMEDICAMENTOS.ReadOnly=True
  VALORPAGOGUIA.ReadOnly=True
  VALORGLOSAGUIA.ReadOnly=True
  VALORTOTALTABELAPROPRIA.ReadOnly=True
  VALORPROCESSADO.ReadOnly=True
  VALORPAGOFORNECEDORES.ReadOnly=True
  TIPOATENDOPERADORAINTERMED.ReadOnly = True
  OPERADORAINTERMEDIARIA.ReadOnly = True
End Sub

Public Sub LiberarCamposParaAjuste()

  Select Case vCodigoCampoComErro
    'Até o momento estes foram os erros tratados para ajuste manual no reenvio do monitoramento
    'Caso ocorra erro em outro campo que não esteja tratado, deverá ser feito a análise do campo pela Benner, para avaliar qual será o tratamento adequado
	Case "012"
      CONTRATADOEXECUTANTECNES.ReadOnly = False
    Case "014"
      CONTRATADOEXECUTANTECPFCNPJ.ReadOnly = False
    Case "015"
      CONTRATADOEXECUTANTEMUNICIPIO.ReadOnly = False
    Case "018"
      BENEFDATANASCIMENTO.ReadOnly = False
    Case "019"
      BENEFMUNICIPIORESIDENCIA.ReadOnly = False
    Case "020"
      BENEFNUMEROREGISTROPLANO.ReadOnly = False
    Case "023"
      NUMEROGUIAPRESTADOR.ReadOnly = False
    Case "026"
      NUMEROGUIASOLICINTERNACAO.ReadOnly = False
    Case "029", "030", "031"
      DATAINICIALFATURAMENTO.ReadOnly = False
      DATAFINALFATURAMENTO.ReadOnly = False
 	Case "035"
	  CBOSEXECUTANTE.ReadOnly = False
    Case "041"
      CIDPRINCIPAL.ReadOnly = False
    Case "042"
      CID2.ReadOnly = False
    Case "043"
      CID3.ReadOnly = False
    Case "044"
      CID4.ReadOnly = False
	Case "045"
      TIPOATENDIMENTO.ReadOnly = False
    Case "081"
      OPERADORAINTERMEDIARIA.ReadOnly = False
	Case "120"
      TIPOATENDOPERADORAINTERMED.ReadOnly = False
  End Select
End Sub

Public Function VerificarSeCampoFoiALterado() As String
  Dim verificaDataRealizacao As Boolean
  Dim registroOriginal As BPesquisa
  Set registroOriginal = NewQuery


  registroOriginal.Add("SELECT *                         ")
  registroOriginal.Add("  FROM ANS_TISMONITORAMENTO_GUIA ")
  registroOriginal.Add(" WHERE HANDLE = :HANDLE          ")

  registroOriginal.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIAORIGEM").AsInteger
  registroOriginal.Active = True

  vCampoFoiALterado = False
  vFoiALteradoDataRealizacao = False
  VerificarSeCampoFoiALterado = ""
  verificaDataRealizacao = True

  If Not registroOriginal.EOF Then

    Select Case vCodigoCampoComErro
	  Case "012"
        If registroOriginal.FieldByName("CONTRATADOEXECUTANTECNES").AsString <> CurrentQuery.FieldByName("CONTRATADOEXECUTANTECNES").AsString Then
          vCampoFoiALterado = True
        End If
      Case "014"
        If registroOriginal.FieldByName("CONTRATADOEXECUTANTECPFCNPJ").AsString <> CurrentQuery.FieldByName("CONTRATADOEXECUTANTECPFCNPJ").AsString Then
          vCampoFoiALterado = True
        End If
      Case "015"
        If registroOriginal.FieldByName("CONTRATADOEXECUTANTEMUNICIPIO").AsString <> CurrentQuery.FieldByName("CONTRATADOEXECUTANTEMUNICIPIO").AsString Then
          vCampoFoiALterado = True
        End If
      Case "018"
        If (registroOriginal.FieldByName("BENEFDATANASCIMENTO").AsDateTime <> CurrentQuery.FieldByName("BENEFDATANASCIMENTO").AsDateTime) Or _
           (registroOriginal.FieldByName("DATAREALIZACAO").AsDateTime <> CurrentQuery.FieldByName("DATAREALIZACAO").AsDateTime) Then

          If CurrentQuery.FieldByName("BENEFDATANASCIMENTO").AsDateTime > CurrentQuery.FieldByName("DATAREALIZACAO").AsDateTime Then
            VerificarSeCampoFoiALterado = "Data de nascimento do beneficiário é maior que a data de realização da guia!"
          End If
          vCampoFoiALterado = True
        End If
        verificaDataRealizacao = False
      Case "019"
        If registroOriginal.FieldByName("BENEFMUNICIPIORESIDENCIA").AsString <> CurrentQuery.FieldByName("BENEFMUNICIPIORESIDENCIA").AsString Then
          vCampoFoiALterado = True
        End If
      Case "020"
        If registroOriginal.FieldByName("BENEFNUMEROREGISTROPLANO").AsString <> CurrentQuery.FieldByName("BENEFNUMEROREGISTROPLANO").AsString Then
          vCampoFoiALterado = True
        End If
      Case "023"
        If registroOriginal.FieldByName("NUMEROGUIAPRESTADOR").AsString <> CurrentQuery.FieldByName("NUMEROGUIAPRESTADOR").AsString Then
          vCampoFoiALterado = True
        End If
      Case "026"
        If registroOriginal.FieldByName("NUMEROGUIASOLICINTERNACAO").AsString <> CurrentQuery.FieldByName("NUMEROGUIASOLICINTERNACAO").AsString Then
          vCampoFoiALterado = True
        End If
      Case "029", "030", "031"

        verificaDataRealizacao = False
        If CurrentQuery.FieldByName("DATAREALIZACAO").AsDateTime > CurrentQuery.FieldByName("DATAPROCESSAMENTO").AsDateTime Then
          VerificarSeCampoFoiALterado = "Data de realização da guia é maior que a data de processamento!"
        End If

        If CurrentQuery.FieldByName("DATAINICIALFATURAMENTO").AsDateTime > CurrentQuery.FieldByName("DATAPROCESSAMENTO").AsDateTime Then
          VerificarSeCampoFoiALterado = "Data inicial do faturamento da guia é maior que a data de processamento!"
        End If

        If CurrentQuery.FieldByName("DATAFINALFATURAMENTO").AsDateTime > CurrentQuery.FieldByName("DATAPROCESSAMENTO").AsDateTime Then
          VerificarSeCampoFoiALterado = "Data final do faturamento da guia é maior que a data de processamento!"
        End If

        If CurrentQuery.FieldByName("DATAFINALFATURAMENTO").AsDateTime < CurrentQuery.FieldByName("DATAINICIALFATURAMENTO").AsDateTime Then
          VerificarSeCampoFoiALterado = "Data final do faturamento da guia é menor que a data inicial de faturamento da guia!"
        End If

        If (registroOriginal.FieldByName("DATAREALIZACAO").AsDateTime <> CurrentQuery.FieldByName("DATAREALIZACAO").AsDateTime)                 Or _
           (registroOriginal.FieldByName("DATAINICIALFATURAMENTO").AsDateTime <> CurrentQuery.FieldByName("DATAINICIALFATURAMENTO").AsDateTime) Or _
           (registroOriginal.FieldByName("DATAFINALFATURAMENTO").AsDateTime <> CurrentQuery.FieldByName("DATAFINALFATURAMENTO").AsDateTime)     Then
          vCampoFoiALterado = True
        End If

	  Case "035"
        If registroOriginal.FieldByName("CBOSEXECUTANTE").AsString <> CurrentQuery.FieldByName("CBOSEXECUTANTE").AsString Then
          vCampoFoiALterado = True
	    End If
      Case "041"
        If registroOriginal.FieldByName("CIDPRINCIPAL").AsInteger <> CurrentQuery.FieldByName("CIDPRINCIPAL").AsInteger Then
          vCampoFoiALterado = True
        End If
      Case "042"
        If registroOriginal.FieldByName("CID2").AsInteger <> CurrentQuery.FieldByName("CID2").AsInteger Then
          vCampoFoiALterado = True
        End If
      Case "043"
        If registroOriginal.FieldByName("CID3").AsInteger <> CurrentQuery.FieldByName("CID3").AsInteger Then
          vCampoFoiALterado = True
        End If
      Case "044"
        If registroOriginal.FieldByName("CID4").AsInteger <> CurrentQuery.FieldByName("CID4").AsInteger Then
          vCampoFoiALterado = True
        End If
  	  Case "045"
	    If registroOriginal.FieldByName("TIPOATENDIMENTO").AsInteger <> CurrentQuery.FieldByName("TIPOATENDIMENTO").AsInteger Then
          vCampoFoiALterado = True
        End If
      Case "081"
	    If registroOriginal.FieldByName("OPERADORAINTERMEDIARIA").AsString <> CurrentQuery.FieldByName("OPERADORAINTERMEDIARIA").AsString Then
          vCampoFoiALterado = True
        End If
	  Case "120"
	    If registroOriginal.FieldByName("TIPOATENDOPERADORAINTERMED").AsInteger <> CurrentQuery.FieldByName("TIPOATENDOPERADORAINTERMED").AsInteger Then
          vCampoFoiALterado = True
        End If
    End Select
  End If

  If verificaDataRealizacao And (registroOriginal.FieldByName("DATAREALIZACAO").AsDateTime <> CurrentQuery.FieldByName("DATAREALIZACAO").AsDateTime) Then
    vFoiALteradoDataRealizacao = True
  End If

  registroOriginal.Active = False
  Set registroOriginal = Nothing

End Function
