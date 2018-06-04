'HASH: 5BC7C2252C982E235662527DD005F7B5
 
'Macro tabela SAM_MATMEDROTREAJUSTE
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
  Dim interface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String


  If VisibleMode Then

    If CurrentQuery.State <>1 Then
      MsgBox("O registro está em edição! Por favor confirme ou cancele as alterações")
      Exit Sub
    End If

    Set interface = CreateBennerObject("BSINTERFACE0044.RotinaMatMed")
    interface.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    CurrentQuery.Active = False
    CurrentQuery.Active = True

    RefreshNodesWithTable("SAM_MATMEDROTREAJUSTE")

  Else
    Set interface = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = interface.ExecucaoImediata(CurrentSystem, _
                                      "SAMREAJUSTEMATMED", _
                                      "RotinaMatMed_CancelaMatMed", _
                                      "Rotina Reajuste MatMed - Cancelamento", _
                                      CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                      "SAM_MATMEDROTREAJUSTE", _
                                      "SITUACAO", _
                                      "", _
                                      "", _
                                      "C", _
                                      False, _
                                      vsMensagem, _
                                      Null)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
    End If
  End If

  Set interface = Nothing

End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim interface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String

  If VisibleMode Then
    If CurrentQuery.State <> 1 Then
      MsgBox("O registro está em edição! Por favor confirme ou cancele as alterações")
      Exit Sub
    End If

    Set interface = CreateBennerObject("BSINTERFACE0044.RotinaMatMed")
    interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    CurrentQuery.Active = False
    CurrentQuery.Active = True

    RefreshNodesWithTable("SAM_MATMEDROTREAJUSTE")

  Else
    Set interface = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = interface.ExecucaoImediata(CurrentSystem, _
                                      "SAMREAJUSTEMATMED", _
                                      "RotinaMatMed_ProcessaMatMed", _
                                      "Rotina Reajuste MatMed - Processamento", _
                                      CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                      "SAM_MATMEDROTREAJUSTE", _
                                      "SITUACAO", _
                                      "", _
                                      "", _
                                      "P", _
                                      False, _
                                      vsMensagem, _
                                      Null)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
    End If
  End If

  Set interface = Nothing

End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOPROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	ElseIf CommandID = "BOTAOCANCELAR" Then
		BOTAOCANCELAR_OnClick
	End If
End Sub
