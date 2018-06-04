'HASH: E0C4E1471884A127D237252E9E81A4C4
'#Uses "*bsShowMessage"

Dim DllInterface As String
Dim NomeProcesso As String
Dim Acao As String

Public Sub ChamadaBotao(Tipo As String)
  Dim HandleRotina As Integer
  HandleRotina = CurrentQuery.FieldByName("CODIGO").AsInteger

  If Tipo = "P" Then
  	DllInterface = "ProcessamentoDissidio"
    NomeProcesso = "Rotina de Dissídio Retroativo ("+CStr(HandleRotina)+")"
  Else
    DllInterface = "CancelamentoDissidio"
    NomeProcesso = "Rotina de Dissídio Retroativo ("+CStr(HandleRotina)+") - Cancelamento"
  End If

  Acao = Tipo
  ExecutarNoServidor(HandleRotina)
  'ExecutarLocal(HandleRotina)

  RefreshNodesWithTable("SAM_ROTINADISSIDIO")
End Sub


Public Sub ConfigurarCampos()
	If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" And Not CurrentQuery.FieldByName("SITUACAO").IsNull Then
		DESCRICAO.ReadOnly = True
		COMPETENCIABASE.ReadOnly = True
		REPLICAR.ReadOnly = True
		CONTRATO.ReadOnly = True
		TIPODEPENDENTE.ReadOnly = True
	Else
		DESCRICAO.ReadOnly = False
		COMPETENCIABASE.ReadOnly = False
		REPLICAR.ReadOnly = False
		CONTRATO.ReadOnly = False
		TIPODEPENDENTE.ReadOnly = False
	End If
End Sub

Public Sub ExecutarNoServidor(HandleRotina As Integer)
   Dim obj As Object
   Dim vsMensagemErro As String
   Dim viRetorno As Long

   Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
   viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                    "Benner.Saude.Processo.DissidioRetroativo", _
                                    DllInterface, _
                                    NomeProcesso, _
                                    HandleRotina, _
                                    "SAM_ROTINADISSIDIO", _
                                    "SITUACAO", _
                                    "", _
                                    "", _
                                    Acao, _
                                    False, _
                                    vsMensagemErro, _
                                    Null)

   If viRetorno = 0 Then
     bsShowMessage("Processo enviado ao servidor, favor verificar o monitor de processos!", "I")
   Else
     bsShowMessage(vsMensagemErro, "I")
   End If


 Set obj = Nothing

End Sub

Public Sub ExecutarLocal(HandleRotina As Integer)
 Dim obj As Object
 If VisibleMode Then

   SessionVar("HANDLE") = CStr(HandleRotina)
   Set obj = CreateBennerObject("Benner.Saude.Processo.DissidioRetroativo." + DllInterface)
   obj.Exec(CurrentSystem)

 End If

 Set obj = Nothing

End Sub

Public Sub BOTAOCANCELAR_OnClick()
	ChamadaBotao("C")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
    ChamadaBotao("P")
End Sub


Public Sub TABLE_AfterScroll()
	ConfigurarCampos
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
		If WebMode Then
			bsShowMessage("Não é possível excluir uma rotina que não esteja em aberto.","I")
		Else
			bsShowMessage("Não é possível excluir uma rotina que não esteja em aberto.","E")
		End If
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	ConfigurarCampos
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	ConfigurarCampos
End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOPROCESSAR"
            ChamadaBotao("P")
        Case "BOTAOCANCELAR"
        	ChamadaBotao("C")
	End Select
End Sub
