'HASH: 6A4BE8A9C4AFF73CE1D87A667F06F86E

'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"

Option Explicit
Dim vCodigoCampoComErro As String
Dim vCodigoGlosa As String
Dim vCampoFoiALterado As Boolean
Dim vHabilitaProcedimento As Boolean

Public Sub PROCEDIMENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, PROCEDIMENTO.Text)
  If vHandle <> 0 Then
    CurrentQuery.FieldByName("PROCEDIMENTO").Value = vHandle
  End If
End Sub

Public Sub TABLE_AfterEdit()
  vHabilitaProcedimento = False

  If SessionVar("CODIGOGLOSA") <> "" Then
    vCodigoGlosa = SessionVar("CODIGOGLOSA")
  End If

  If SessionVar("CODIGOCAMPOCOMERRO") <> "" Then
    vCodigoCampoComErro = SessionVar("CODIGOCAMPOCOMERRO")
    LiberarCamposParaAjuste
  End If
End Sub

Public Sub TABLE_AfterPost()
  If vCampoFoiALterado Then
    Dim component As CSBusinessComponent
    Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Procedimento.Correcao, Benner.Saude.ANS.Processos")
    component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    component.AddParameter(pdtString, vCodigoCampoComErro)
    component.Execute("RemoverErro")
    Set component = Nothing

    RefreshNodesWithTable("ANS_TISMONITORAMENTO_GUIA_PROC")
  End If
End Sub

Public Sub TABLE_AfterScroll()
  BloquearTodosCampos

  If WebMode Then
    CODIGOTABELA.WebLocalWhere = "A.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
  Else
    CODIGOTABELA.LocalWhere = "VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If SessionVar("CODIGOCAMPOCOMERRO") <> "" Then
    VerificarSeCampoFoiALterado

    If Not vCampoFoiALterado Then
      bsShowMessage("Não foi realizado alteração no campo com erro","E")
      CanContinue = False
    Else
      SessionVar("CODIGOCAMPOCOMERRO") = ""
      SessionVar("CODIGOGLOSA") = ""
      vCodigoGlosa = ""
    End If
  End If
End Sub

Public Sub BloquearTodosCampos
	CODIGOTABELA.ReadOnly = True
	GRUPOPROCEDIMENTO.ReadOnly = True
	PROCEDIMENTO.ReadOnly = True
	DENTE.ReadOnly = True
	REGIAO.ReadOnly = True
	DENTEFACE.ReadOnly = True
	QUANTIDADEINFORMADA.ReadOnly = True
	QUANTIDADEPAGA.ReadOnly = True
	VALORINFORMADO.ReadOnly = True
	VALORPAGO.ReadOnly = True
	VALORPAGOFORNECEDOR.ReadOnly = True
	CNPJFORNECEDOR.ReadOnly = True
	TABNDENTEREGIAO.ReadOnly = True
	VALORPF.ReadOnly = True
End Sub

Public Sub LiberarCamposParaAjuste()
  vHabilitaProcedimento = vCodigoCampoComErro = "066" And vCodigoGlosa = "2601"
  Select Case vCodigoCampoComErro
    'Até o momento estes foram os erros tratados para ajuste manual no reenvio do monitoramento
    'Caso ocorra erro em outro campo que não esteja tratado, deverá ser feito a análise do campo pela Benner, para avaliar qual será o tratamento adequado
	Case "064", "065", "066"
      CODIGOTABELA.ReadOnly = False
      If vHabilitaProcedimento Then
		PROCEDIMENTO.ReadOnly = False
      End If
    Case "067", "068", "069"
      DENTE.ReadOnly = False
	  REGIAO.ReadOnly = False
	  DENTEFACE.ReadOnly = False
	  TABNDENTEREGIAO.ReadOnly = False
    Case "070"
      QUANTIDADEINFORMADA.ReadOnly = False
    Case "071"
      VALORINFORMADO.ReadOnly = False
    Case "072"
      QUANTIDADEPAGA.ReadOnly = False
    Case "073"
      VALORPAGO.ReadOnly = False
  End Select
End Sub


Public Function VerificarSeCampoFoiALterado()

  Dim registroOriginal As BPesquisa
  Set registroOriginal = NewQuery

  registroOriginal.Add("SELECT *                              ")
  registroOriginal.Add("  FROM ANS_TISMONITORAMENTO_GUIA_PROC ")
  registroOriginal.Add(" WHERE HANDLE = :HANDLE               ")

  registroOriginal.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PROCEDIMENTOORIGEM").AsInteger
  registroOriginal.Active = True

  vCampoFoiALterado = False

  If Not registroOriginal.EOF Then
    Select Case vCodigoCampoComErro
	    Case "064", "065", "066"
	        If registroOriginal.FieldByName("CODIGOTABELA").AsInteger <> CurrentQuery.FieldByName("CODIGOTABELA").AsInteger Then
	          vCampoFoiALterado = True
	        End If
	        If vHabilitaProcedimento And registroOriginal.FieldByName("PROCEDIMENTO").AsInteger <> CurrentQuery.FieldByName("PROCEDIMENTO").AsInteger Then
			  vCampoFoiALterado = True
	        End If
        Case "067", "068", "069"
	        If registroOriginal.FieldByName("DENTE").AsString <> CurrentQuery.FieldByName("DENTE").AsString Then
	          vCampoFoiALterado = True
	        End If
	        If registroOriginal.FieldByName("REGIAO").AsString <> CurrentQuery.FieldByName("REGIAO").AsString Then
	          vCampoFoiALterado = True
	        End If
	         If registroOriginal.FieldByName("DENTEFACE").AsString <> CurrentQuery.FieldByName("DENTEFACE").AsString Then
	          vCampoFoiALterado = True
	        End If
            If registroOriginal.FieldByName("TABNDENTEREGIAO").AsString <> CurrentQuery.FieldByName("TABNDENTEREGIAO").AsString Then
	          vCampoFoiALterado = True
	        End If
        Case "070"
	        If registroOriginal.FieldByName("QUANTIDADEINFORMADA").AsFloat <> CurrentQuery.FieldByName("QUANTIDADEINFORMADA").AsFloat Then
	          vCampoFoiALterado = True
	        End If
        Case "071"
	        If registroOriginal.FieldByName("VALORINFORMADO").AsFloat <> CurrentQuery.FieldByName("VALORINFORMADO").AsFloat Then
	          vCampoFoiALterado = True
	        End If
        Case "072"
	        If registroOriginal.FieldByName("QUANTIDADEPAGA").AsFloat <> CurrentQuery.FieldByName("QUANTIDADEPAGA").AsFloat Then
	          vCampoFoiALterado = True
	        End If
        Case "073"
	        If registroOriginal.FieldByName("VALORPAGO").AsFloat <> CurrentQuery.FieldByName("VALORPAGO").AsFloat Then
	          vCampoFoiALterado = True
	        End If
    End Select
  End If

  registroOriginal.Active = False
  Set registroOriginal = Nothing

End Function

Public Sub table_OnSavebtnClick(CanContinue As Boolean)

  If CurrentQuery.FieldByName("VALORPF").IsNull Then
	CurrentQuery.FieldByName("VALORPF").AsInteger = 0
  End If

End Sub
