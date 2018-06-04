'HASH: 07D4D0AED59C7AB138E5EA35B8A25CF0
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()

  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "5" Then
	bsShowMessage("Processo abortado. A rotina não está Processada!","I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOXML").AsString <> "1" Then
    bsShowMessage("Processo abortado. Já foi gerado o XML para esta rotina!","I")
    Exit Sub
  End If

  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Cancelando a rotina da competência: " + CurrentQuery.FieldByName("COMPETENCIA").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.CancelaMonitoramento"
  processo.SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  processo.Execute

  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAO").AsString = "2"
  CurrentQuery.Post

  Set processo = Nothing

  bsShowMessage("Processo enviado para execução no servidor!","I")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOEXCLUIRERROSGUIA_OnClick()

 If Not WebMode Then
    If bsShowMessage("Deseja excluir todos os erros das guias e procedimentos?", "Q") = vbNo Then
      Exit Sub
    End If
  End If

  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
	bsShowMessage("Processo abortado. A rotina não está aberta!","I")
	Exit Sub
  End If


  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Excluindo erros das guias/procedimentos, da competência: " + CurrentQuery.FieldByName("COMPETENCIA").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.ExcluirErros"
  processo.SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  processo.Execute


  Set processo = Nothing

  bsShowMessage("Processo enviado para execução no servidor!","I")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOGERARREENVIO_OnClick()
  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAORETORNO").AsString <> "5" Then
	bsShowMessage("Processo abortado. A rotina não está com o retorno processado!","I")
	Exit Sub
  End If

  Dim processo As CSServerExec
  Set processo = NewServerExec
  'Dim processo As Object
  'Set processo = CreateBennerObject("Benner.Saude.ANS.Processos.ProcessoGerarReenvio")
  'SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  'processo.Exec(CurrentSystem)

  processo.Description = "Monitoramento TISS - Gerando rotina de reenvio da competência: " + CurrentQuery.FieldByName("COMPETENCIA").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.ProcessoGerarReenvio"
  processo.SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  processo.Execute

  Set processo = Nothing

  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAORETORNO").AsString = "4"
  CurrentQuery.Post

  bsShowMessage("Processo enviado para execução no servidor!","I")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOINDICARPROTOCOLOPTA_OnClick()
  Dim form As CSVirtualForm
  Dim vsMensagem As String

  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOXML").AsString <> "5" Then
    bsShowMessage("Processo abortado. O arquivo XML ainda não foi gerado para essa rotina!","I")
    Exit Sub
  End If

  SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString

  If VisibleMode Then
    Set form = NewVirtualForm

    form.Caption = "TISS Monitoramento - Indicar protocolo PTA"
    form.TableName = "TV_MONITORAMENTO_INDICACAOPTA"
    form.Show
    Set form = Nothing
  End If

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
	bsShowMessage("Processo abortado. A rotina não está mais aberta!","I")
	Exit Sub
  End If

  Dim sql As Object
  Set sql = NewQuery

  If CurrentQuery.FieldByName("TABREENVIO").AsInteger = 2 Then
    sql.Add("SELECT 1                             ")
    sql.Add("  FROM ANS_TISMONITORAMENTO_GUIA     ")
    sql.Add(" WHERE ROTINAMONITORAMENTO = :ROTINA ")
    sql.Add("   AND SITUACAO <> '2'               ")
    sql.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.Active = True

    If Not sql.EOF Then
      If bsShowMessage("Processo abortado. Existe guia na rotina de reenvio que não foi corrigida!", "I") Then
        Set sql = Nothing
        Exit Sub
      End If
    End If
  End If

  sql.Active = False
  sql.Clear
  sql.Add("SELECT 1 QTDE                     ")
  sql.Add("  FROM ANS_TISMONITORAMENTO       ")
  sql.Add(" WHERE HANDLE <> :HANDLE          ")
  sql.Add("   AND OPERADORA = :OPERADORA     ")
  sql.Add("   AND COMPETENCIA = :COMPETENCIA ")
  sql.Add("   AND SITUACAO <> '1'            ")
  sql.Add("   AND SITUACAO <> '5'            ")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ParamByName("OPERADORA").AsInteger = CurrentQuery.FieldByName("OPERADORA").AsInteger
  sql.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  sql.Active = True

  If Not sql.EOF Then
    bsShowMessage("Processo abortado. Já existe outra rotina da mesma competência e operadora sendo processada!","I")
	Exit Sub
  End If

  Set sql = Nothing


  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Processando a rotina da competência: " + CurrentQuery.FieldByName("COMPETENCIA").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.ProcessaMonitoramento"
  processo.SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  processo.Execute

  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAO").AsString = "2"
  CurrentQuery.Post

  'Dim processo As Object
  'Set processo = CreateBennerObject("Benner.Saude.ANS.Processos.ProcessaMonitoramento")
  'SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  'processo.Exec(CurrentSystem)

  Set processo = Nothing

  bsShowMessage("Processo enviado para execução no servidor!","I")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOPROCESSARRETORNO_OnClick()

  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOXML").AsString <> "5" Then
	bsShowMessage("Processo abortado. O XML desta rotina ainda não está processado!","I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("PROTOCOLOPTA").AsString = "" Then
    bsShowMessage("Processo abortado. Não foi indicado o protocolo PTA para esta rotina!","I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAORETORNO").AsString <> "1" Then
	bsShowMessage("Processo abortado. Já foi processado o retorno do XML para esta rotina!","I")
	Exit Sub
  End If

  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Processando retorno da competência: " + CurrentQuery.FieldByName("COMPETENCIA").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.ProcessaRetornoXml"
  processo.SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  processo.Execute

  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAORETORNO").AsString = "2"
  CurrentQuery.Post

  Set processo = Nothing

  bsShowMessage("Processo enviado para execução no servidor!","I")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub CANCELARXML_OnClick()

  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOXML").AsString <> "5" Then
	bsShowMessage("Processo abortado. O XML desta rotina ainda não está processado!","I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("PROTOCOLOPTA").AsString <> "" Then
    bsShowMessage("Processo abortado. Já foi indicado o protocolo PTA para esta rotina!","I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAORETORNO").AsString <> "1" Then
    bsShowMessage("Processo abortado. Já foi processado o retorno do XML para esta rotina!","I")
    Exit Sub
  End If

  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Cancelando o XML da competência: " + CurrentQuery.FieldByName("COMPETENCIA").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.CancelaXmlEnvio"
  processo.SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  processo.Execute

  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAOXML").AsString = "2"
  CurrentQuery.Post

  Set processo = Nothing

  bsShowMessage("Processo enviado para execução no servidor!","I")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub GERARXML_OnClick()

  If CurrentQuery.State <>1 Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "5" Then
	bsShowMessage("Processo abortado. A rotina não está Processada!","I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOXML").AsString <> "1" Then
	bsShowMessage("Processo abortado. Já foi gerado o XML para esta rotina!","I")
	Exit Sub
  End If

  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Gerando o XML da competência: " + CurrentQuery.FieldByName("COMPETENCIA").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.GeraXmlEnvio"
  processo.SessionVar("HANDLE_ROTMONITORAMENTOTISS") = CurrentQuery.FieldByName("HANDLE").AsString
  processo.Execute

  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAOXML").AsString = "2"
  CurrentQuery.Post

  Set processo = Nothing

  bsShowMessage("Processo enviado para execução no servidor!","I")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub TABLE_AfterScroll()
  If Not WebMode Then
    BOTAOEXCLUIRERROSGUIA.Visible = False
    If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
	  BOTAOPROCESSAR.Visible = True
	  BOTAOCANCELAR.Visible = False
      BOTAOINDICARPROTOCOLOPTA.Visible = False
      BOTAOPROCESSARRETORNO.Visible = False
      GERARXML.Visible = False
      CANCELARXML.Visible = False
      BOTAOGERARREENVIO.Visible = False
      If CurrentQuery.FieldByName("TABREENVIO").AsInteger = 2 Then
         BOTAOEXCLUIRERROSGUIA.Visible = True
      End If
    Else
      BOTAOPROCESSAR.Visible = False
      BOTAOGERARREENVIO.Visible = False

      If CurrentQuery.FieldByName("SITUACAO").AsString = "5" And CurrentQuery.FieldByName("SITUACAOXML").AsString = "1" Then
	    BOTAOCANCELAR.Visible = True
	    GERARXML.Visible = True
	    CANCELARXML.Visible = False
        BOTAOINDICARPROTOCOLOPTA.Visible = False
        BOTAOPROCESSARRETORNO.Visible = False
      Else
        BOTAOCANCELAR.Visible = False
        GERARXML.Visible = False

        If CurrentQuery.FieldByName("SITUACAOXML").AsString = "5" And CurrentQuery.FieldByName("SITUACAORETORNO").AsString = "1" Then
          BOTAOINDICARPROTOCOLOPTA.Visible = True

          If CurrentQuery.FieldByName("PROTOCOLOPTA").AsString <> "" Then
            BOTAOPROCESSARRETORNO.Visible = True
            CANCELARXML.Visible = False
          Else
            BOTAOPROCESSARRETORNO.Visible = False
            CANCELARXML.Visible = True
          End If
	    Else
	      CANCELARXML.Visible = False
          BOTAOINDICARPROTOCOLOPTA.Visible = False
          BOTAOPROCESSARRETORNO.Visible = False
          If CurrentQuery.FieldByName("SITUACAORETORNO").AsString = "5" And ExibeBotaoReenvio Then
            BOTAOGERARREENVIO.Visible = True
          End If
	    End If
      End If
    End If
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TABREENVIO").AsInteger = 2 Then
    If VisibleMode Then
      If bsShowMessage("Deseja excluir todos os registros da rotina de reenvio?", "Q") = vbYes Then
        ExcluirGuiasDoReenvio
      Else
        CanContinue = False
        Exit Sub
      End If
    Else
      ExcluirGuiasDoReenvio
    End If
  End If
End Sub
Public Function ExcluirGuiasDoReenvio
  Dim component As CSBusinessComponent
  Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Rotina.ExcluirGuiasReenvio, Benner.Saude.ANS.Processos")
  component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  component.Execute("Excluir")
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
	Case "BOTAOPROCESSAR"
	  BOTAOPROCESSAR_OnClick
	Case "BOTAOCANCELAR"
	  BOTAOCANCELAR_OnClick
	Case "BOTAOINDICARPROTOCOLOPTA"
	  BOTAOINDICARPROTOCOLOPTA_OnClick
	Case "BOTAOPROCESSARRETORNO"
	  BOTAOPROCESSARRETORNO_OnClick
	Case "CANCELARXML"
	  CANCELARXML_OnClick
	Case "GERARXML"
	  GERARXML_OnClick
	Case "BOTAOGERARREENVIO"
	  BOTAOGERARREENVIO_OnClick
	Case "BOTAOEXCLUIRERROSGUIA"
	  BOTAOEXCLUIRERROSGUIA_OnClick
  End Select
End Sub

Public Sub TABLE_OnDeleteBtnClick(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    bsShowMessage("Não é possível excluir o registro, pois ele já foi processado!","I")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Function ExibeBotaoReenvio As Boolean
  Dim qSql As BPesquisa
  Set qSql = NewQuery

  qSql.Add("SELECT 1                     ")
  qSql.Add("  FROM ANS_TISMONITORAMENTO  ")
  qSql.Add(" WHERE ENVIOORIGEM = :ROTINA ")

  qSql.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSql.Active = True

  ExibeBotaoReenvio = False

  If qSql.EOF Then
    qSql.Active = False
    qSql.Clear
    qSql.Add("SELECT 1                             ")
    qSql.Add("  FROM ANS_TISMONITORAMENTO_GUIA     ")
    qSql.Add(" WHERE ROTINAMONITORAMENTO = :ROTINA ")
    qSql.Add("   AND SITUACAO = '4'                ")

    qSql.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSql.Active = True

    If qSql.EOF Then
      qSql.Active = False
      qSql.Clear
      qSql.Add("SELECT 1                             ")
      qSql.Add("  FROM ANS_TISMONITORAMENTO_ARQUIVO  ")
      qSql.Add(" WHERE ROTINAMONITORAMENTO = :ROTINA ")
      qSql.Add("   AND TABRESULTADORETORNO = 2       ")

      qSql.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qSql.Active = True

      If Not qSql.EOF Then
        ExibeBotaoReenvio = True
      End If
    Else
      ExibeBotaoReenvio = True
    End If
  End If

  qSql.Active = False
  Set qSql = Nothing
End Function
