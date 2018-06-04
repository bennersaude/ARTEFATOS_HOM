'HASH: BE670360C787388A5F88C1CE17F45E90
Option Explicit
'#Uses "*bsShowMessage"

Public Sub BOTAOPROCESSAR_OnClick()

'Dim sx As Object
Dim sx As CSServerExec
Set sx = NewServerExec


If (CurrentQuery.State = 2) Then
  bsShowMessage("O registro está em edição.", "I")
  Exit Sub
End If

If (CurrentQuery.FieldByName("PEG").IsNull) Then
  bsShowMessage("Salvar o PEG antes de processar.", "I")
  Exit Sub
End If

If (CurrentQuery.FieldByName("SITUACAO").Value = 1) Then
  CurrentQuery.Edit
  CurrentQuery.FieldByName("SITUACAO").AsString = "9"
  CurrentQuery.Post
End If

sx.Description = "Processamento da Rotina de Importação de Ressarcimento ao SUS"
sx.DllClassName = "BENNER.SAUDE.SERVICES.PROCCONTAS.RESSARCIMENTOSUS.ImportacaoRessarcimentoSUS"
sx.SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
sx.Execute

'Set sx = CreateBennerObject("BENNER.SAUDE.SERVICES.PROCCONTAS.RESSARCIMENTOSUS.ImportacaoRessarcimentoSUS")
'SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
'sx.Exec(CurrentSystem)

Set sx = Nothing

bsShowMessage("Processo enviado para execução no servidor!", "I")

If VisibleMode Then
  RefreshNodesWithTable("SAM_ROTRESSARCIMENTOSUS")
End If

End Sub

Public Sub CarregarXml()
	Dim sx As CSServerExec
    Set sx = NewServerExec
    'Dim sx As Object

    sx.Description = "Importação do arquivo XML"
    sx.DllClassName = "BENNER.SAUDE.SERVICES.PROCCONTAS.RESSARCIMENTOSUS.ImportacaoXml"
    sx.SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
    sx.Execute

    'Set sx = CreateBennerObject("BENNER.SAUDE.SERVICES.PROCCONTAS.RESSARCIMENTOSUS.ImportacaoXml")
    'SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
    'sx.Exec(CurrentSystem)

    Set sx = Nothing

    bsShowMessage("Processo enviado para execução no servidor!", "I")
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
  Case "BOTAOCARREGARXML"
    CarregarXml
  Case "BOTAOPROCESSAR"
  	BOTAOPROCESSAR_OnClick
  End Select

  CanContinue = True
End Sub
