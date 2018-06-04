'HASH: BAC5BC0E490EA222F3CE66BB38067F02
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterPost()
  'Incluir registro processo
  Dim component As CSBusinessComponent
  Dim handleProcesso As Long

  Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Reenvio.CriarRotinaAjuste, Benner.Saude.ANS.Processos")
  component.AddParameter(pdtInteger, RecordHandleOfTable("ANS_TISMONITORAMENTO"))
  component.AddParameter(pdtInteger, CLng(SessionVar("HANDLECAMPOMONITORAMENTO")))

  handleProcesso = CLng(component.Execute("Criar"))

  Set component = Nothing

  'chamar agendamento
  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Ajustando os procedimentos com o erro ""065"""
  processo.DllClassName = "Benner.Saude.ANS.Processos.ReprocessarCampo065"
  processo.SessionVar("HANDLE_ROTINAREAJUSTE") = CStr(handleProcesso)
  processo.SessionVar("HANDLE_GRUPOEVENTO") = CurrentQuery.FieldByName("GRUPOEVENTO").AsString
  processo.SessionVar("HANDLE_CODIGOTABELA") = CurrentQuery.FieldByName("CODIGOTABELA").AsString

  bsShowMessage("Processo enviado para execução no servidor. Verifique o andamento na carga ""Status reprocessamento"" da rotina do monitoramento.","I")

  processo.Execute

  Set processo = Nothing

End Sub
