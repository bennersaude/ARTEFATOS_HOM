'HASH: 75AE678E7DDC35ADFD215C0FFAF4ED34
'#Uses "*bsShowMessage"
Dim handleRecebedor As Long

Public Sub TABLE_AfterScroll()
  Dim qConsulta As BPesquisa
  Set qConsulta = NewQuery
  qConsulta.Clear

  qConsulta.Add("SELECT IDENTIFICADORPAGAMENTO, RECEBEDOR FROM SAM_PEG WHERE HANDLE = :HANDLE ")
  qConsulta.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_PEG")
  qConsulta.Active = True

  If(CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
    CurrentQuery.FieldByName("IDENTIFICADORPAGAMENTO").AsString = qConsulta.FieldByName("IDENTIFICADORPAGAMENTO").AsString
  End If

  handleRecebedor = qConsulta.FieldByName("RECEBEDOR").AsInteger

  Set qConsulta   = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim DLLEspecifico As Object
  Dim vsMsgVerifica As String

  Set DLLEspecifico = CreateBennerObject("ESPECIFICO.UESPECIFICO")
  vbResultado = DLLEspecifico.PRO_VerificaIdentificadorPagamento(CurrentSystem, RecordHandleOfTable("SAM_PEG"), handleRecebedor, CurrentQuery.FieldByName("IDENTIFICADORPAGAMENTO").AsString, vsMsgVerifica)
  Set DLLEspecifico = Nothing

  If (vbResultado) Then
    bsshowmessage(vsMsgVerifica, "E")
    CanContinue = False
    Exit Sub
  End If

  Dim AlteracaoPeg As CSBusinessComponent

  Set AlteracaoPeg = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.SamPegBLL, Benner.Saude.ProcessamentoContas.Business")
  AlteracaoPeg.AddParameter(pdtInteger, RecordHandleOfTable("SAM_PEG"))
  AlteracaoPeg.AddParameter(pdtString, CurrentQuery.FieldByName("IDENTIFICADORPAGAMENTO").AsString)
  AlteracaoPeg.Execute("AlterarIdentificadorPagamento")

  Set AlteracaoPeg = Nothing
End Sub
