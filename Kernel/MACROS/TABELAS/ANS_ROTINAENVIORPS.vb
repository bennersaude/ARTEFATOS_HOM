'HASH: 39332C7AD2DA7D09F9550C56B667A4A6
 '#Uses "*bsShowMessage"


Public Sub BOTAOCANCELAR_OnClick()

  If Not RegistroExistenteFisicamente Then
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOROTINA").AsInteger <> 5 Then
    bsShowMessage("A rotina não pode ser cancelada, pois a situação da rotina não está Processada", "I")
    Exit Sub
  End If

  Dim qJaExisteRotinaAberta As Object
  Set qJaExisteRotinaAberta = NewQuery
  qJaExisteRotinaAberta.Clear
  qJaExisteRotinaAberta.Add("SELECT COUNT(1) QTD            ")
  qJaExisteRotinaAberta.Add("  FROM ANS_ROTINAENVIORPS      ")
  qJaExisteRotinaAberta.Add(" WHERE HANDLE <> :HANDLE       ")
  qJaExisteRotinaAberta.Add("   and OPERADORA =  :OPERADORA ")
  qJaExisteRotinaAberta.Add("   and SITUACAOROTINA <>  '5'  ")
  qJaExisteRotinaAberta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qJaExisteRotinaAberta.ParamByName("OPERADORA").AsInteger = CurrentQuery.FieldByName("OPERADORA").AsInteger
  qJaExisteRotinaAberta.Active = True
  If qJaExisteRotinaAberta.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Impossível cancelar, pois já existe uma rotina na situação de aberto para a operadora em questão, favor verificar", "I")
    Exit Sub
  End If
  Set qJaExisteRotinaAberta = Nothing

  Dim qJaExisteRotinaPosteriorProcessada As Object
  Set qJaExisteRotinaPosteriorProcessada = NewQuery
  qJaExisteRotinaPosteriorProcessada.Clear
  qJaExisteRotinaPosteriorProcessada.Add("SELECT COUNT(1) QTD                            ")
  qJaExisteRotinaPosteriorProcessada.Add("  FROM ANS_ROTINAENVIORPS                      ")
  qJaExisteRotinaPosteriorProcessada.Add(" WHERE HANDLE <> :HANDLE                       ")
  qJaExisteRotinaPosteriorProcessada.Add("   and OPERADORA =  :OPERADORA                 ")
  qJaExisteRotinaPosteriorProcessada.Add("   and SITUACAOROTINA =  '5'                   ")
  qJaExisteRotinaPosteriorProcessada.Add("   and DATAHORAPROCESSAMENTO > :DATAHORAROTINA ")
  qJaExisteRotinaPosteriorProcessada.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qJaExisteRotinaPosteriorProcessada.ParamByName("OPERADORA").AsInteger = CurrentQuery.FieldByName("OPERADORA").AsInteger
  qJaExisteRotinaPosteriorProcessada.ParamByName("DATAHORAROTINA").AsDateTime = CurrentQuery.FieldByName("DATAHORAPROCESSAMENTO").AsDateTime
  qJaExisteRotinaPosteriorProcessada.Active = True
  If CurrentQuery.FieldByName("ENVIADOANS").AsString = "S" Then
    bsShowMessage("Impossível cancelar, rotina já enviada para ANS", "I")
    Exit Sub
  End If
  Set qJaExisteRotinaPosteriorProcessada = Nothing


  Dim qSetaRotinaAberta As Object
  Set qSetaRotinaAberta = NewQuery
  qSetaRotinaAberta.Clear
  qSetaRotinaAberta.Add("UPDATE ANS_ROTINAENVIORPS                ")
  qSetaRotinaAberta.Add("   SET SITUACAOROTINA = '1',             ")
  qSetaRotinaAberta.Add("       USUARIOCANCELAMENTO = :USUARIO,   ")
  qSetaRotinaAberta.Add("       DATAHORACANCELAMENTO = :DATAHORA, ")
  qSetaRotinaAberta.Add("       SITUACAOGERACAOXML = '1',         ")
  qSetaRotinaAberta.Add("       USUARIOPROCESSAMENTO = NULL,      ")
  qSetaRotinaAberta.Add("       DATAHORAPROCESSAMENTO = NULL      ")
  qSetaRotinaAberta.Add(" WHERE HANDLE = :HANDLE                  ")
  qSetaRotinaAberta.ParamByName("USUARIO").AsInteger = CurrentUser
  qSetaRotinaAberta.ParamByName("DATAHORA").AsDateTime = ServerNow
  qSetaRotinaAberta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSetaRotinaAberta.ExecSQL
  Set qSetaRotinaAberta = Nothing

  ClearFieldDocument("ANS_ROTINAENVIORPS", "ARQUIVOINCLUSAO", CurrentQuery.FieldByName("HANDLE").AsInteger, True)
  ClearFieldDocument("ANS_ROTINAENVIORPS", "ARQUIVOALTERACAO", CurrentQuery.FieldByName("HANDLE").AsInteger, True)
  ClearFieldDocument("ANS_ROTINAENVIORPS", "ARQUIVOEXCLUSAO", CurrentQuery.FieldByName("HANDLE").AsInteger, True)
  ClearFieldDocument("ANS_ROTINAENVIORPS", "ARQUIVOVINCULACAO", CurrentQuery.FieldByName("HANDLE").AsInteger, True)

  bsShowMessage("Cancelamento executado com sucesso", "I")
  RefreshNodesWithTable("ANS_ROTINAENVIORPS")

End Sub

Public Sub BOTAOEXCLUIRDADOS_OnClick()

  If Not RegistroExistenteFisicamente Then
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOROTINA").AsInteger <> 1 Then
    bsShowMessage("Dados não podem ser excluídos, pois a situação da rotina não está Aberta", "I")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("SITUACAOGERACAODADOS").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOGERACAODADOS").AsInteger <> 5) Then
    bsShowMessage("Dados não podem ser excluídos, pois ainda existem geração de dados a serem realizados", "I")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 5) Then
    bsShowMessage("Rotina ainda em processamento de exclusão dos dados", "I")
    Exit Sub
  End If

  Dim qRotinaSemRegistros As Object
  Set qRotinaSemRegistros = NewQuery
  qRotinaSemRegistros.Clear
  qRotinaSemRegistros.Add("SELECT COUNT(1) QTD                                   ")
  qRotinaSemRegistros.Add("  FROM ANS_ROTINAENVIORPS_PREST RP                    ")
  qRotinaSemRegistros.Add(" WHERE RP.ROTINAENVIORPS = :ROTINAENVIORPS            ")
  qRotinaSemRegistros.ParamByName("ROTINAENVIORPS").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qRotinaSemRegistros.Active = True
  If qRotinaSemRegistros.FieldByName("QTD").AsInteger = 0 Then
    bsShowMessage("Rotina sem dado a ser excluído, favor verificar", "I")
    Exit Sub
  End If
  Set qRotinaSemRegistros = Nothing

  If bsShowMessage("Deseja excluir todos os registros de prestadores da rotina?", "Q") = vbNo Then
    Exit Sub
  End If

  Dim sx As CSServerExec
  Set sx = NewServerExec

  sx.Description = "RPS - Exclusão de dados"
  sx.DllClassName = "Benner.Saude.ANS.Processos.ExcluirDadosRPS"

  sx.SessionVar("RPS_HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString

  Dim qRotinaAgendada As Object
  Set qRotinaAgendada = NewQuery
  qRotinaAgendada.Clear
  qRotinaAgendada.Add("UPDATE ANS_ROTINAENVIORPS              ")
  qRotinaAgendada.Add("   SET SITUACAOEXCLUSAODADOS = '2'     ")
  qRotinaAgendada.Add(" WHERE HANDLE = :HANDLE")
  qRotinaAgendada.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qRotinaAgendada.ExecSQL
  Set qRotinaAgendada = Nothing

  sx.Execute
  Set sx = Nothing

  bsShowMessage("Processo enviado para o servidor", "I")

End Sub

Public Sub BOTAOGERARDADOS_OnClick()

  If Not RegistroExistenteFisicamente Then
    Exit Sub
  End If

  If VisibleMode Then
    Dim INTERFACE0002 As Object
    Dim vsMensagem As String
    Dim vcContainer As CSDContainer

    SessionVar("RPS_HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString

    Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

    INTERFACE0002.Exec(CurrentSystem, _
                       1, _
                       "TV_FORM0089", _
                       "Gerar dados RPS", _
                       0, _
                       270, _
                       500, _
                       False, _
                       vsMensagem, _
                       vcContainer)

    Set INTERFACE0002 = Nothing
  End If
End Sub

Public Sub BOTAOGERARXML_OnClick()

  If Not RegistroExistenteFisicamente Then
    Exit Sub
  End If

If CurrentQuery.FieldByName("SITUACAOGERACAOXML").AsInteger <> 1 Then
    bsShowMessage("XML não pode ser gerado, pois a situação da rotina não está Aberta", "I")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("SITUACAOGERACAODADOS").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOGERACAODADOS").AsInteger <> 5) Then
    bsShowMessage("XML não pode ser gerado, pois ainda existem geração de dados a serem realizados", "I")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 5) Then
    bsShowMessage("Rotina ainda em processamento de exclusão dos dados", "I")
    Exit Sub
  End If

  Dim sx As CSServerExec
  Set sx = NewServerExec

  sx.Description = "RPS - Geração do XML"
  sx.DllClassName = "Benner.Saude.ANS.Processos.GerarXMLRPS"

  sx.SessionVar("RPS_HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString

  Dim qRotinaAgendada As Object
  Set qRotinaAgendada = NewQuery
  qRotinaAgendada.Clear
  qRotinaAgendada.Add("UPDATE ANS_ROTINAENVIORPS       ")
  qRotinaAgendada.Add("   SET SITUACAOGERACAOXML = '2' ")
  qRotinaAgendada.Add(" WHERE HANDLE = :HANDLE         ")
  qRotinaAgendada.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qRotinaAgendada.ExecSQL

  sx.Execute
  Set sx = Nothing

  bsShowMessage("Processo enviado para o servidor", "I")

End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If Not RegistroExistenteFisicamente Then
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOROTINA").AsInteger <> 1 Then
    bsShowMessage("A rotina não pode ser processada, pois a situação da rotina não está Aberta", "I")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("SITUACAOGERACAODADOS").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOGERACAODADOS").AsInteger <> 5) Then
    bsShowMessage("A rotina não pode ser processada, pois ainda existem geração de dados a serem realizados", "I")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 5) Then
    bsShowMessage("A rotina não pode ser processada, pois ainda existem registros a serem excluídos", "I")
    Exit Sub
  End If


  Dim qRotinaSemRegistros As Object
  Set qRotinaSemRegistros = NewQuery
  qRotinaSemRegistros.Clear
  qRotinaSemRegistros.Add("SELECT COUNT(1) QTD                                   ")
  qRotinaSemRegistros.Add("  FROM ANS_ROTINAENVIORPS_PREST RP                    ")
  qRotinaSemRegistros.Add(" WHERE RP.ROTINAENVIORPS = :ROTINAENVIORPS            ")
  qRotinaSemRegistros.ParamByName("ROTINAENVIORPS").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qRotinaSemRegistros.Active = True
  If qRotinaSemRegistros.FieldByName("QTD").AsInteger = 0 Then
    bsShowMessage("Rotina sem prestador para ser processado, favor verificar", "I")
    Exit Sub
  End If
  Set qRotinaSemRegistros = Nothing


  Dim qRegistrosSemVinculo As Object
  Set qRegistrosSemVinculo = NewQuery
  qRegistrosSemVinculo.Clear
  qRegistrosSemVinculo.Add("SELECT COUNT(1) QTD                                       ")
  qRegistrosSemVinculo.Add("  FROM ANS_ROTINAENVIORPS_PREST RP                        ")
  qRegistrosSemVinculo.Add(" WHERE RP.ROTINAENVIORPS = :ROTINAENVIORPS                ")
  qRegistrosSemVinculo.Add("   AND NOT EXISTS (SELECT 1                               ")
  qRegistrosSemVinculo.Add("                     FROM ANS_ROTINAENVIORPS_PREST_VINC   ")
  qRegistrosSemVinculo.Add("                    WHERE ROTINAENVIORPSPREST = RP.HANDLE)")
  qRegistrosSemVinculo.ParamByName("ROTINAENVIORPS").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qRegistrosSemVinculo.Active = True
  If qRegistrosSemVinculo.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Existem prestadores sem vínculo do produto, favor verificar", "I")
    Exit Sub
  End If
  Set qRegistrosSemVinculo = Nothing

  Dim qSetaRotinaProcessada As Object
  Set qSetaRotinaProcessada = NewQuery
  qSetaRotinaProcessada.Clear
  qSetaRotinaProcessada.Add("UPDATE ANS_ROTINAENVIORPS                ")
  qSetaRotinaProcessada.Add("   SET SITUACAOROTINA = '5',             ")
  qSetaRotinaProcessada.Add("       USUARIOCANCELAMENTO = NULL,       ")
  qSetaRotinaProcessada.Add("       DATAHORACANCELAMENTO = NULL,      ")
  qSetaRotinaProcessada.Add("       USUARIOPROCESSAMENTO = :USUARIO,  ")
  qSetaRotinaProcessada.Add("       DATAHORAPROCESSAMENTO = :DATAHORA ")
  qSetaRotinaProcessada.Add(" WHERE HANDLE = :HANDLE                  ")
  qSetaRotinaProcessada.ParamByName("USUARIO").AsInteger = CurrentUser
  qSetaRotinaProcessada.ParamByName("DATAHORA").AsDateTime = ServerNow
  qSetaRotinaProcessada.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSetaRotinaProcessada.ExecSQL

  Set qSetaRotinaProcessada = Nothing

  bsShowMessage("Processamento executado com sucesso", "I")
  RefreshNodesWithTable("ANS_ROTINAENVIORPS")

End Sub

Public Sub BOTAOENVIARANS_OnClick()
	Dim qSetaEnvioAns As Object
	Set qSetaEnvioAns = NewQuery
	If CurrentQuery.FieldByName("ENVIADOANS").AsString <> "S" Then
		If bsShowMessage("Deseja confirmar o envio do arquivo a ANS?", "Q") = vbYes Then
			qSetaEnvioAns.Clear
			qSetaEnvioAns.Add("UPDATE ANS_ROTINAENVIORPS    ")
			qSetaEnvioAns.Add("   SET ENVIADOANS = 'S'      ")
			qSetaEnvioAns.Add(" WHERE HANDLE = :HANDLE      ")
			qSetaEnvioAns.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			qSetaEnvioAns.ExecSQL

		End If
	Else
		If bsShowMessage("Deseja cancelar a confirmação de envio do arquivo?", "Q") = vbYes Then

			qSetaEnvioAns.Clear
			qSetaEnvioAns.Add("UPDATE ANS_ROTINAENVIORPS    ")
			qSetaEnvioAns.Add("   SET ENVIADOANS = 'N'      ")
			qSetaEnvioAns.Add(" WHERE HANDLE = :HANDLE      ")
			qSetaEnvioAns.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			qSetaEnvioAns.ExecSQL

		End If
	End If

	RefreshNodesWithTable("ANS_ROTINAENVIORPS")

End Sub

Public Sub TABLE_AfterScroll()

	If CurrentQuery.FieldByName("ENVIADOANS").AsString = "S" Then
		BOTAOENVIARANS.Hint = "Cancelar confirmação de envio do arquivo para a ANS"
		BOTAOENVIARANS.Caption = ("Cancelar Envio")
	Else
		BOTAOENVIARANS.Hint = "Realiza a confirmação de que pelo menos um arquivo da rotina foi enviado à ANS"
		BOTAOENVIARANS.Caption = ("Confirmar envio")
	End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim qChecaRotinaComDado As Object
  Set qChecaRotinaComDado = NewQuery
  qChecaRotinaComDado.Clear
  qChecaRotinaComDado.Add("SELECT COUNT(1) QTD                                   ")
  qChecaRotinaComDado.Add("  FROM ANS_ROTINAENVIORPS_PREST RP                    ")
  qChecaRotinaComDado.Add(" WHERE RP.ROTINAENVIORPS = :ROTINAENVIORPS            ")
  qChecaRotinaComDado.ParamByName("ROTINAENVIORPS").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qChecaRotinaComDado.Active = True
  If qChecaRotinaComDado.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Rotina com dados gerados, favor verificar", "E")
    CanContinue = False
    Exit Sub
  End If
  Set qChecaRotinaComDado = Nothing

  Dim qChecaRotinaProduto As Object
  Set qChecaRotinaProduto = NewQuery
  qChecaRotinaProduto.Clear
  qChecaRotinaProduto.Add("SELECT COUNT(1) QTD                                   ")
  qChecaRotinaProduto.Add("  FROM ANS_ROTINAENVIORPS_PRODUTO RP                    ")
  qChecaRotinaProduto.Add(" WHERE RP.ROTINAENVIORPS = :ROTINAENVIORPS            ")
  qChecaRotinaProduto.ParamByName("ROTINAENVIORPS").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qChecaRotinaProduto.Active = True
  If qChecaRotinaProduto.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Rotina com produtos vinculados, favor verificar", "E")
    CanContinue = False
    Exit Sub
  End If
  Set qChecaRotinaProduto = Nothing

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("SITUACAOROTINA").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOROTINA").AsInteger <> 5) Then
    bsShowMessage("Processamento da rotina em andamento, favor verificar", "E")
    CanContinue = False
    Exit Sub
  End If
  If (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 5) Then
    bsShowMessage("Exclusão dos dados em andamento, favor verificar", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qExisteOutraRotina As Object
  Set qExisteOutraRotina = NewQuery
  qExisteOutraRotina.Clear
  qExisteOutraRotina.Add("SELECT HANDLE ")
  qExisteOutraRotina.Add("  FROM ANS_ROTINAENVIORPS ")
  qExisteOutraRotina.Add(" WHERE HANDLE <> :HANDLE")
  qExisteOutraRotina.Add("  AND OPERADORA = :OPERADORA")
  qExisteOutraRotina.Add("  AND SITUACAOROTINA <> :SITUACAOROTINA")
  qExisteOutraRotina.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qExisteOutraRotina.ParamByName("OPERADORA").AsInteger = CurrentQuery.FieldByName("OPERADORA").AsInteger
  qExisteOutraRotina.ParamByName("SITUACAOROTINA").AsString = "5"
  qExisteOutraRotina.Active = True
  If qExisteOutraRotina.FieldByName("HANDLE").AsInteger > 0 Then
      bsShowMessage("Existe outra rotina para a operadora escolhida que ainda não está processada", "E")
      CanContinue = False
      Exit Sub
  End If

  Set qExisteOutraRotina = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
 	Case "BOTAOCANCELAR"
 	  BOTAOCANCELAR_OnClick
 	Case "BOTAOGERARDADOS"
 	  BOTAOGERARDADOS_OnClick
 	Case "BOTAOEXCLUIRDADOS"
      BOTAOEXCLUIRDADOS_OnClick
 	Case "BOTAOPROCESSAR"
 	  BOTAOPROCESSAR_OnClick
 	Case "BOTAOGERARXML"
 	  BOTAOGERARXML_OnClick
 	Case "BOTAOENVIARANS"
 	  BOTAOENVIARANS_OnClick
  End Select
End Sub
Public Function RegistroExistenteFisicamente() As Boolean
  RegistroExistenteFisicamente = True
  If CurrentQuery.State = 3 Then
    bsShowMessage("Registro deve ser salvo antes de processar qualquer funcionalidade!", "I")
    RegistroExistenteFisicamente = False
  End If
End Function
