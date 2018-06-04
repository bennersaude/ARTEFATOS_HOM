'HASH: A2A6570A307322409B68C831239998DD

'#Uses "*bsShowMessage"
Option Explicit

Public Sub BOTAOREPROCESSAR_OnClick()
  If CurrentQuery.FieldByName("CODIGO").AsString = "065" Or CurrentQuery.FieldByName("CODIGO").AsString = "066" _
     Or CurrentQuery.FieldByName("CODIGO").AsString = "064" Or CurrentQuery.FieldByName("CODIGO").AsString = "067" _
     Or CurrentQuery.FieldByName("CODIGO").AsString = "068" Or CurrentQuery.FieldByName("CODIGO").AsString = "069" Then

    If VisibleMode Then
      SessionVar("HANDLECAMPOMONITORAMENTO") = CurrentQuery.FieldByName("HANDLE").AsString
      SessionVar("CAMPOMONITORAMENTOPROCESSAR")=CurrentQuery.FieldByName("CODIGO").AsString

      If CurrentQuery.FieldByName("CODIGO").AsString = "065" Then
        Reprocessar065
      Else
        Reprocessar066 'vai executar quando fo 066 e 064
      End If
    End If
  Else
    ReprocessarDemaisCampos
  End If
End Sub

Public Sub ReprocessarDemaisCampos()
  'Incluir registro processo
  Dim component As CSBusinessComponent
  Dim handleProcesso As Long

  Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Reenvio.CriarRotinaAjuste, Benner.Saude.ANS.Processos") ' formato: [namespace.classe], [assembly]
  component.AddParameter(pdtInteger, RecordHandleOfTable("ANS_TISMONITORAMENTO"))
  component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)

  handleProcesso = CLng(component.Execute("Criar"))

  Set component = Nothing

  'chamar agendamento
  Dim processo As CSServerExec
  Set processo = NewServerExec

  processo.Description = "Monitoramento TISS - Ajustando os procedimentos com o erro " + CurrentQuery.FieldByName("CODIGO").AsString
  processo.DllClassName = "Benner.Saude.ANS.Processos.ReprocessarCampos"
  processo.SessionVar("HANDLE_ROTINAREAJUSTE") = CStr(handleProcesso)
  processo.Execute

  'Dim processo As Object
  'Set processo = CreateBennerObject("Benner.Saude.ANS.Processos.ReprocessarCampos")
  'SessionVar("HANDLE_ROTINAREAJUSTE") = CStr(handleProcesso)
  'processo.Exec(CurrentSystem)

  Set processo = Nothing

  bsShowMessage("Processo enviado para execução no servidor. Verifique o andamento na carga ""Status reprocessamento"" da rotina do monitoramento.","I")
End Sub

Public Sub Reprocessar065()
  Dim form As CSVirtualForm
  Set form = NewVirtualForm

  form.Caption = "Ajuste dos procedimentos com erro"
  form.TableName = "TV_MONITORAMENTO_AJUSTE_GRUPO"
  form.Height = 150
  form.Width = 400
  form.Show

  Set form = Nothing
End Sub

Public Sub Reprocessar066()
  Dim form As CSVirtualForm
  Set form = NewVirtualForm

  form.Caption = "Ajuste dos grupos de procedimento com erro"
  form.TableName = "TV_MONITORAMENTO_AJUSTE_EVENTO"
  form.Height = 200
  form.Width = 500
  form.Show

  Set form = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
	Case "BOTAOREPROCESSAR"
	  BOTAOREPROCESSAR_OnClick
	Case "REPROCESSAARCAMPO065","REPROCESSAARCAMPO066"
	  SessionVar("HANDLECAMPOMONITORAMENTO") = CurrentQuery.FieldByName("HANDLE").AsString
  End Select
End Sub
