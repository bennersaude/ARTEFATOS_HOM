'HASH: 8F775B10B3505A65610EC05A8F0ACBFB
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub TABLE_AfterPost()
  'Incluir registro processo
  Dim component As CSBusinessComponent
  Dim handleProcesso As Long
  Dim nomeDLL As String

  Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Reenvio.CriarRotinaAjuste, Benner.Saude.ANS.Processos")
  component.AddParameter(pdtInteger, RecordHandleOfTable("ANS_TISMONITORAMENTO"))
  component.AddParameter(pdtInteger, CLng(SessionVar("HANDLECAMPOMONITORAMENTO")))

  handleProcesso = CLng(component.Execute("Criar"))

  nomeDLL="Benner.Saude.ANS.Processos.ReprocessarCampo066"

  If SessionVar("CAMPOMONITORAMENTOPROCESSAR")="067" _
    Or SessionVar("CAMPOMONITORAMENTOPROCESSAR")="068" _
    Or SessionVar("CAMPOMONITORAMENTOPROCESSAR")="069" Then

    nomeDLL= "Benner.Saude.ANS.Processos.ReprocessarCampos067e068e069"
  End If

  Set component = Nothing

  'chamar agendamento
  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Ajustando os procedimentos com o erro "+SessionVar("CAMPOMONITORAMENTOPROCESSAR")
  processo.DllClassName = nomeDLL
  processo.SessionVar("HANDLE_ROTINAREAJUSTE") = CStr(handleProcesso)
  processo.SessionVar("TABAJUSTE") = CurrentQuery.FieldByName("TABAJUSTE").AsString
  processo.SessionVar("HANDLE_EVENTO") = CurrentQuery.FieldByName("EVENTO").AsString
  processo.SessionVar("HANDLE_CODIGOTABELA") = CurrentQuery.FieldByName("CODIGOTABELA").AsString

  bsShowMessage("Processo enviado para execução no servidor. Verifique o andamento na carga ""Status reprocessamento"" da rotina do monitoramento.","I")

  processo.Execute

  Set processo = Nothing


End Sub


Public Sub TABLE_AfterScroll()
 If Not WebMode Then
   If (SessionVar("CAMPOMONITORAMENTOPROCESSAR")="067" _
    Or SessionVar("CAMPOMONITORAMENTOPROCESSAR")="068" _
    Or SessionVar("CAMPOMONITORAMENTOPROCESSAR")="069") _
    And (CurrentQuery.State <> 1) Then
    CurrentQuery.FieldByName("TABAJUSTE").AsInteger = 1
  End If
 End If
End Sub
