'HASH: E435A4B187DA949A89FD733D70F8FBC4
'#Uses "*bsShowMessage"
Public Sub BOTAOGERAR_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    bsShowMessage("Não é possível gerar! Situação está pendente","I")
    Exit Sub
  End If



  Dim CONFIRMA As Object
  Set CONFIRMA = NewQuery

  Dim viRetornoMensagem As Long

  CONFIRMA.Active = False
  CONFIRMA.Clear
  CONFIRMA.Add("Select B.SITUACAO                    ")
  CONFIRMA.Add("  FROM SAM_ROTINAIMP_BENEF  B,       ")
  CONFIRMA.Add("       SAM_ROTINAIMP_FAM    F,       ")
  CONFIRMA.Add("       SAM_ROTINAIMP_FILIAL FL       ")
  CONFIRMA.Add(" WHERE B.IMPFAM          = F.HANDLE  ")
  CONFIRMA.Add("   And F.ROTINAIMPFILIAL = FL.HANDLE ")
  CONFIRMA.Add("   And B.SITUACAO        = 'H'       ")
  CONFIRMA.Add("   And FL.HANDLE         = :HANDLE   ")
  CONFIRMA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  CONFIRMA.Active = True

  If CONFIRMA.EOF Then
    Dim qRotina As Object
    Set qRotina = NewQuery
    Dim IMPORTA As Object
    Dim vsMensagemErro As String

    qRotina.Active = False
    qRotina.Clear
    qRotina.Add("SELECT TABTIPOIMPORTACAO")
    qRotina.Add("  FROM SAM_ROTINAIMP ")
    qRotina.Add(" WHERE HANDLE = :ROTINAIMP")
    qRotina.ParamByName("ROTINAIMP").AsInteger = CurrentQuery.FieldByName("ROTINAIMP").AsInteger
    qRotina.Active = True

    If qRotina.FieldByName("TABTIPOIMPORTACAO").AsInteger = 3 Then
      If VisibleMode Then
        Set IMPORTA = CreateBennerObject("BSINTERFACE0015.RotinasImportacaoBenef")
        viRetornoMensagem = IMPORTA.Confirmar(CurrentSystem, CurrentQuery.FieldByName("ROTINAIMP").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
      Else


		Dim vcContainer As CSDContainer
	   	Set vcContainer = NewContainer
	   	vcContainer.AddFields("HFILIAL:INTEGER; HANDLE:INTEGER")

	    vcContainer.Insert
	    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAIMP").AsInteger
	    vcContainer.Field("HFILIAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger


		Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")

        viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                     "BSBEN015", _
                                                     "ImportarConfirmar", _
                                                     "Importação de beneficiários - Confirmar", _
                                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                     "SAM_ROTINAIMP_FILIAL", _
                                                     "SITUACAOGERAR", _
                                                     "", _
                                                     "", _
                                                     "P", _
                                                     True, _
                                                     vsMensagemErro, _
                                                     vcContainer)

        If viRetorno = 0 Then
           bsShowMessage("Processo enviado para execução no servidor!", "I")
        Else
           bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
        End If

      End If

      If viRetornoMensagem = 1 Then
        bsShowMessage("Ocorreu erro no processo","I")
      End If

    Else
      If VisibleMode Then
        Set IMPORTA = CreateBennerObject("BSInterface0025.RotinasImportacaoBenef")
        viRetornoMensagem = IMPORTA.Confirmar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ROTINAIMP").AsInteger)
      Else
        Set IMPORTA = CreateBennerObject("BSBEN005.RotinaImportar_Confirmar")
        viRetornoMensagem = IMPORTA.Exec(CurrentSystem, CurrentQuery.FieldByName("ROTINAIMP").AsInteger, vsMensagemErro, 0, CurrentQuery.FieldByName("HANDLE").AsInteger)
      End If

      If viRetornoMensagem = 1 Then
        bsShowMessage("Ocorreu erro no processo:" + Chr(13) + vsMensagemErro,"I")
      End If


    End If
    Set IMPORTA = Nothing

  Else
    bsShowMessage("ATENÇÃO! Ainda existem Beneficiários homônimos pendentes, impossível continuar!","I")
  End If

  Dim qryAdesao As Object
  Set qryAdesao = NewQuery
  qryAdesao.Active = False
  qryAdesao.Clear
  qryAdesao.Add("Select COUNT(B.HANDLE) QTD           ")
  qryAdesao.Add("  FROM SAM_ROTINAIMP_BENEF  B,       ")
  qryAdesao.Add("       SAM_ROTINAIMP_FAM    F,       ")
  qryAdesao.Add("       SAM_ROTINAIMP_FILIAL FL,      ")
  qryAdesao.Add("       SAM_CONTRATO C                ")
  qryAdesao.Add(" WHERE B.IMPFAM          = F.HANDLE  ")
  qryAdesao.Add("   And F.ROTINAIMPFILIAL = FL.HANDLE ")
  qryAdesao.Add("   And FL.HANDLE         = :HANDLE   ")
  qryAdesao.Add("   AND F.CONTRATO = C.HANDLE         ")
  qryAdesao.Add("   AND B.DATAADESAO < C.DATAADESAO   ")
  qryAdesao.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qryAdesao.Active = True

  If qryAdesao.FieldByName("QTD").AsInteger >0 Then
    bsShowMessage("Existe beneficiário com data de adesão inferior a data de adesão do contrato", "I")
  End If

  Set qryAdesao = Nothing
  Set CONFIRMA = Nothing

  'Inicio SMS 81946 - Rodrigo Andrade - 01/08/2007
  'Criado parametro para verificar a necessidade de executar essa procedure.

  Dim qParamBenef As Object
  Set qParamBenef = NewQuery
  qParamBenef.Active = False
  qParamBenef.Clear
  qParamBenef.Add("SELECT CONSISTEHISTORICO FROM SAM_PARAMETROSBENEFICIARIO")
  qParamBenef.Active = True

  If qParamBenef.FieldByName("CONSISTEHISTORICO").AsString = "S" Then
    'Lopes - sms 49746 - Chamar a procedure para corrigir o historico dos beneficiarios

' SMS 92341 - Paulo Melo - 22/02/2008 - Erro de FK, pois era passado 1 como handle do motivo de cancelamento padrão e não havia handle = 1 na tabela do cliente
    Dim handlemotivo As String          ' agora é passado o motivo com menor handle na tabela SAM_MOTIVOCANCELAMENTO
	Dim q As Object
	Set q = NewQuery

	q.Add("SELECT MIN(HANDLE) MOTIVO")
	q.Add("FROM SAM_MOTIVOCANCELAMENTO")
	q.Active = True

	handlemotivo = q.FieldByName("MOTIVO").AsString

    Dim res As String
	Dim dll As Object
	Set dll = CreateBennerObject("SAMPROCEDURE.UIProcedure")
'	dll.ExecProc(CurrentSystem, "BSCORRIGEBENEFCANCELADOS", "pHndMotivoCancPadrao;1;I;I|", res)
	dll.ExecProc(CurrentSystem, "BSCORRIGEBENEFCANCELADOS", "pHndMotivoCancPadrao;"+handlemotivo+";I;I|", res)
' SMS 92341 - Paulo Melo - 22/02/2008 - FIM
  End If
  qParamBenef.Active = False
  Set qParamBenef = Nothing
  Set dll= Nothing
  'FIm SMS 81946
  RefreshNodesWithTable("SAM_ROTINAIMP_FILIAL")

End Sub

Public Sub TABLE_AfterScroll()
  Dim sql As Object
  Dim sqlup As Object
  Dim SITUACAO As String
  Set sqlup = NewQuery
  Set sql = NewQuery

  SITUACAO = CurrentQuery.FieldByName("SITUACAO").AsString

  sql.Active = False
  sql.Add("SELECT A.*")
  sql.Add("  FROM SAM_ROTINAIMP_FAM A")
  sql.Add(" WHERE ROTINAIMPFILIAL =:IMPFILIAL")
  sql.Add("   AND (A.ERRO = 'S'")
  sql.Add("        OR EXISTS (SELECT B.HANDLE")
  sql.Add("                     FROM SAM_ROTINAIMP_BENEF B")
  sql.Add("                    WHERE B.IMPFAM = A.HANDLE")
  sql.Add("                      AND ((B.SITUACAO <> 'G') and (B.SITUACAO <> 'P') and (B.SITUACAO <> 'R') and (B.SITUACAO <> 'H'))  ))")
  sql.ParamByName("IMPFILIAL").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If sql.EOF Then
    sqlup.Active = False
    sqlup.Clear
    sqlup.Add("UPDATE SAM_ROTINAIMP_FILIAL SET SITUACAO = 'O' ")
    sqlup.Add(" WHERE HANDLE = :HANDLE")
    sqlup.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    sqlup.ExecSQL
    SITUACAO = "O"
  End If


  If SITUACAO = "P" Then
    BOTAOGERAR.Enabled = False
  Else
    BOTAOGERAR.Enabled = True
  End If

  Set sql = Nothing
  Set sqlup = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
    Case "BOTAOGERAR"
      BOTAOGERAR_OnClick
  End Select
End Sub
