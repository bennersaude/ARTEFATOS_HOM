'HASH: 03AD51E75664FC1A7D82251B01F25056
'Macro: SFN_ROTINAFINREVISAOIRRF

'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim SQL As Object
  Dim vsMensagem As String

  Set SQL = NewQuery

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "5" Then
    bsShowMessage("A Rotina não está cancelada!", "E")
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0054.RevisaoIRRF_Cancelamento")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagem)

  Else
    Dim viRetorno As Long
    Dim vcContainer As CSDContainer
    Set vcContainer = NewContainer

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SfnRecolhimento", _
                                     "Rotina_RevisaoIRRF_Cancelamento", _
                                     "Rotina de Cancelamento de Revisão de IRRF", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINREVISAOIRRF", _
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
   Set Obj = Nothing

  RefreshNodesWithTable("SFN_ROTINAFINREVISAOIRRF")

End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
   bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim SQL As Object
  Dim vsMensagem As String

  Set SQL = NewQuery

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    bsShowMessage("A Rotina não está aberta!", "E")
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0054.RevisaoIRRF_Processamento")
    Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagem)

  Else
    Dim viRetorno As Long

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SfnRecolhimento", _
                                     "Rotina_RevisaoIRRF_Processamento", _
                                     "Rotina de Revisão de IRRF", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINREVISAOIRRF", _
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
   Set Obj = Nothing

  RefreshNodesWithTable("SFN_ROTINAFINREVISAOIRRF")

End Sub

Public Sub TABLE_AfterScroll()

  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
    BOTAOPROCESSAR.Enabled = True
  Else
    BOTAOPROCESSAR.Enabled = False
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "5" Then
    BOTAOCANCELAR.Enabled  = True
  Else
    BOTAOCANCELAR.Enabled  = False
  End If

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
		CanContinue = False
        bsShowMessage("A Rotina não está aberta!", "E")
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
		CanContinue = False
        bsShowMessage("A Rotina não está aberta!", "E")
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOPROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	ElseIf CommandID = "BOTAOCANCELAR" Then
		BOTAOCANCELAR_OnClick
	End If
End Sub


