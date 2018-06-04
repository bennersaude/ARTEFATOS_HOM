'HASH: 53790A9DC943A3B84DB36FD12CE78858
'Macro: SAM_PLANO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOCENTROCUSTO_OnClick()
  Dim Interface As Object

  If VisibleMode Then
    Set Interface = CreateBennerObject("Financeiro.Geral")
    Interface.PlanoCentroCusto(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Interface = Nothing
  Else
    Dim vsMensagemErro As String
    Dim viRetorno As Long
    Dim vcContainer As CSDContainer

    Set vcContainer = NewContainer
    vcContainer.AddFields("HPLANO:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HPLANO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    Set Interface = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Interface.ExecucaoImediata(CurrentSystem, _
	                                "Financeiro", _
	                                "PlanoCentroCusto_Exec", _
	                                "Atualização Centro de Custo nos contratos: " + _
	                                CurrentQuery.FieldByName("DESCRICAO").AsString, _
	                                0, _
	                                "", _
	                                "", _
	                                "", _
	                                "", _
	                                "P", _
	                                False, _
	                                vsMensagemErro, _
	                                vcContainer)
    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If
End Sub

Public Sub TABLE_AfterScroll()
  'Daniela -SMS 12220 -Convênio no registro da ANS
  If Not CurrentQuery.FieldByName("CONVENIO").IsNull Then
    CONVENIO.ReadOnly = True
  Else
    CONVENIO.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If(Not CurrentQuery.FieldByName("DATAVALIDADE").IsNull)And _
     (CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime <CurrentQuery.FieldByName("DATACRIACAO").AsDateTime)Then
  bsShowMessage("A Data de validade, se informada, deve ser maior ou igual a criação", "E")
  CanContinue = False
Else
  CanContinue = True
End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
  	Case "BOTAOCENTROCUSTO"
      BOTAOCENTROCUSTO_OnClick
  End Select
End Sub
