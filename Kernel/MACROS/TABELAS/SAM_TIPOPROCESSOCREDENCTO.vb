'HASH: 9E05C9B854048225CB8BEF6AF79DA548
'Macro: SAM_TIPOPROCESSOCREDENCTO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOSELECIONARRELATORIO_OnClick()
  Dim OLEAutorizador As Object
  Dim lista(4)As String
  lista(1) = "Deferido"
  lista(2) = "Indeferido"
  lista(3) = "Em Análise"
  lista(4) = "Devolvido"

  Begin Dialog UserDialog 250, 161 ' %GRID:10,7,1,1
    OKButton 30, 133, 90, 21
    CancelButton 130, 133, 90, 21
    Text 20, 7, 220, 14, "Selecione o Relatório Desejado", .Text1
    OptionGroup.Group1
    OptionButton 30, 28, 190, 14, "Deferido", .OptionButton1
    OptionButton 30, 49, 190, 14, "Indeferido", .OptionButton2
    OptionButton 30, 70, 190, 14, "Em Análise", .OptionButton3
    OptionButton 30, 91, 190, 14, "Devolvido", .OptionButton4
    OptionButton 30, 112, 190, 14, "Parcialmente Deferido", .OPtionButton5
    End Dialog

    Dim dlg As UserDialog
	Dim Handlexx As Long

    On Error GoTo cancel
    Dialog dlg

    Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
    Select Case dlg.Group1
      Case 0
        Handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório|Código", "UPPER(NOME) LIKE '% DEFERIDO%'", "Procura por Relatórios", True, "")
      Case 1
        Handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório|Código", "UPPER(NOME) LIKE '% INDEFERIDO%'", "Procura por Relatórios", True, "")
      Case 2
        Handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório|Código", "UPPER(NOME) LIKE '% ANALISE%'", "Procura por Relatórios", True, "")
      Case 3
        Handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório|Código", "UPPER(NOME) LIKE '% DEVOLVIDO%'", "Procura por Relatórios", True, "")
      Case 4
        Handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório|Código", "UPPER(NOME) LIKE '% PARCIAL%'", "Procura por Relatórios", True, "")
    End Select

    If Handlexx <> 0 Then
      CurrentQuery.Edit
      Select Case dlg.Group1
        Case 0
          CurrentQuery.FieldByName("HANDLERELATORIODEFERIDO").Value = Handlexx
        Case 1
          CurrentQuery.FieldByName("HANDLERELATORIOINDEFERIDO").Value = Handlexx
        Case 2
          CurrentQuery.FieldByName("HANDLERELATORIOEMANALISE").Value = Handlexx
        Case 3
          CurrentQuery.FieldByName("HANDLERELATORIODEVOLVIDO").Value = Handlexx
        Case 4
          CurrentQuery.FieldByName("HANDLERELATORIOPARCIAL").Value = Handlexx
      End Select
      CurrentQuery.Post
    End If
   cancel :
  Set OLEAutorizador = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim SamPrestadorProcBLL As CSBusinessComponent
  Dim retorno As Boolean
  Set SamPrestadorProcBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcBLL, Benner.Saude.Prestadores.Business")
  SamPrestadorProcBLL.AddParameter(pdtString, "CREDENCIAMENTOAVANCADO")
  GRUPOCONTROLESADICIONAIS.Visible = SamPrestadorProcBLL.Execute("VerificarParametrosParaCredenciamentoAutomatico")
  Set SamPrestadorProcBLL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If((CurrentQuery.FieldByName("VALIDARDOCUMENTOSAOFINALIZAR").AsString = "S") And (CurrentQuery.FieldByName("CONTROLARDOCUMENTACAO").AsString = "N")) Then
	bsShowMessage("Para validar documentos ao finalizar é obrigatório ativar o controle de documentos exigidos.", "E")
	CanContinue = False
  End If

  Dim terminaCom As String
  If(Not CurrentQuery.FieldByName("MODELOTERMO").IsNull) Then
  	terminaCom =  LCase(CurrentQuery.FieldByName("MODELOTERMO").Value)
  	If(terminaCom Like "*.rtf") Then
	  CanContinue = True
	Else
      bsShowMessage("Arquivo de Modelo de Termo de Credenciamento deve ter o formato/extensão RTF!", "E")
      CurrentQuery.FieldByName("MODELOTERMO").Value = Null
    End If
  End If
End Sub
