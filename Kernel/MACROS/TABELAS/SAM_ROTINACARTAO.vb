'HASH: 60AA62484FFEFFDF4F34D50F1ACA3B78
'#Uses "*bsShowMessage"
'Macro: SAM_ROTINACARTAO
'#Uses "*ProcuraContrato"

'A funcao NodeInternalCode é utilizada para determinar se a carga correspondente é da Tarefas de Modelo,
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina
'Alteração: 26/12/2005
'      SMS: 52120 - Marcelo Barbosa

Option Explicit

Dim VPAR As Boolean

Public Sub BOTAOAGENDAR_OnClick()
  Dim qr As Object
  Dim qr1 As Object
  Dim vSituacao As String
  Dim vTabela As String
  Dim vLegendaAgendamento As String
  Dim VLegendaAberta As String
  Dim VLegendaProcessada As String
  Set qr = NewQuery
  Set qr1 = NewQuery
  vTabela = "SAM_ROTINACARTAO"
  vLegendaAgendamento = "3"
  VLegendaAberta = "1"
  VLegendaProcessada = "5"
  qr.Clear
  qr.Add("SELECT SITUACAO FROM " + vTabela + " WHERE HANDLE = :pHANDLE")
  qr.ParamByName("pHandle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qr.Active = True
  vSituacao = qr.FieldByName("SITUACAO").AsString
  If vSituacao <> vLegendaAgendamento Then
    If (vSituacao = VLegendaAberta And CurrentQuery.FieldByName("DATAGERACAO").IsNull) Or _
        (vSituacao = VLegendaProcessada And CurrentQuery.FieldByName("DATAFATURAR").IsNull) Then
      If bsShowMessage("Confirme o agendamento da rotina", "Q") = vbYes Then '(6=yes, 7=não)
        qr1.Clear
        If Not InTransaction Then StartTransaction
        qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
        qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qr1.ParamByName("pSituacao").AsString = vLegendaAgendamento
        qr1.ExecSQL
        If InTransaction Then Commit
      End If
    Else
      bsShowMessage("Rotina já foi faturada.", "I")
    End If
  Else
    If bsShowMessage("Rotina já está agendada. Para retirar o agendamento pressione 'SIM'", "Q") = vbYes Then
      qr1.Clear
      If Not InTransaction Then StartTransaction
      qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
      qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      If (CurrentQuery.FieldByName("DATAGERACAO").IsNull) Then
        qr1.ParamByName("pSituacao").AsString = VLegendaAberta
      Else
        qr1.ParamByName("pSituacao").AsString = VLegendaProcessada
      End If
      qr1.ExecSQL
      If InTransaction Then Commit
    End If
  End If
  Set qr = Nothing
  Set qr1 = Nothing
  RefreshNodesWithTable("SAM_ROTINACARTAO")

End Sub

Public Sub BOTAOCANCELAFATURAMENTO_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAFATURAR").IsNull Then
    bsShowMessage("O faturamento ainda não foi processada", "I")
    Exit Sub
  End If

  'If Not CurrentQuery.FieldByName("DATACANCFATURAMENTO").IsNull Then
  '	 MsgBox("O cancelamento já processado.")
  '	 Exit Sub
  'End If

  If VisibleMode Then
       Set Obj = CreateBennerObject("SAMROTINACARTAO.Geral")
       Obj.CancelarFatura(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
       Set Obj = Nothing

       WriteAudit("C", HandleOfTable("SAM_ROTINACARTAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Cartões - Cancelamento das faturas")
  Else
        Dim vsMensagemErro As String
   		Dim viRetorno As Long

        Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_CancelarFatura", _
                                     "Rotina de Fatura de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO", _
                                     "SITUACAOCANCELAFATURA", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                      vsMensagemErro, _
                                     Null)


        If viRetorno = 0 Then
  			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
    		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      	End If

  End If

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
    bsShowMessage("A Geração ainda não foi processada", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATACANCELAR").IsNull Then
    bsShowMessage("Falta data do cancelamento", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("MOTIVOCANCELAR").IsNull Then
    bsShowMessage("Falta motivo do cancelamento", "I")
    Exit Sub
  End If

  If VisibleMode Then

    Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")
    Obj.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Else

    Dim vsMensagemErro As String
    Dim viRetorno As Long

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_Cancelar", _
                                     "Rotina de Cancelamento de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO", _
                                     "SITUACAOCANCELAMENTO", _
                                     "", _
                                     "", _
                                     "P", _
                                     True, _
                                      vsMensagemErro, _
                                     Null)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

  End If


  Set Obj = Nothing


  WriteAudit("C", HandleOfTable("SAM_ROTINACARTAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Cartões - Cancelamento")

  CurrentQuery.Active = False
  CurrentQuery.Active = True

End Sub

Public Sub BOTAOCOMUNICADO_OnClick()

  'inicio
  Dim lista(2)As String
  Dim vMessage As String
  Dim FormFiltroTexto As String
  Dim TextoFinal As String
  Dim QtdCartao As Long
  Dim RelatorioComunicadoHandle As Long
  Dim QueryDadosRotinaCartao As Object
  Dim QueryBuscaRelatorioComunicadoHandle As Object 'busca o handle do relatório de comunicado de ingresso
  Dim QueryContaQtdComunicado As Object

  Set QueryDadosRotinaCartao = NewQuery
  Set QueryBuscaRelatorioComunicadoHandle = NewQuery
  Set QueryContaQtdComunicado = NewQuery

  vMessage = "================================================== " + Chr(13)

  UserParam = 0
  UserParam = CurrentQuery.FieldByName("HANDLE").AsInteger


  QueryDadosRotinaCartao.Add("SELECT										")
  QueryDadosRotinaCartao.Add("	RC.CODIGO		ROTINACARTAO_CODIGO,		")
  QueryDadosRotinaCartao.Add("	RC.DESCRICAO	ROTINACARTAO_DESCRICAO,")
  QueryDadosRotinaCartao.Add("	RC.DATAROTINA	ROTINACARTAO_DATAROTINA, ")
  QueryDadosRotinaCartao.Add("	RC.OCORRENCIAS	OCORRENCIAS")
  QueryDadosRotinaCartao.Add("FROM SAM_ROTINACARTAO RC")
  QueryDadosRotinaCartao.Add("WHERE RC.HANDLE=:P_ROTINACARTAO_HANDLE")

  QueryBuscaRelatorioComunicadoHandle.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = :P_CODIGO")



  QueryDadosRotinaCartao.ParamByName("P_ROTINACARTAO_HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  QueryDadosRotinaCartao.Active = False
  QueryDadosRotinaCartao.Active = True

  'lista(1)="Comunicado de Ingresso."
  'lista(2)="Comunicado de Renovação."


  Begin Dialog UserDialog 470, 231 ' %GRID:10,7,1,1
    OKButton 130, 77, 90, 21
    CancelButton 240, 77, 90, 21
    Text 20, 7, 340, 14, "Selecione o Tipo de Comunicado a ser Gerado:", .Text1
    OptionGroup.Group1
    OptionButton 110, 28, 240, 14, "Comunicado de Ingresso.", .OptionButton1
    OptionButton 110, 49, 250, 14, "Comunicado de Renovação.", .OptionButton2
    TextBox 10, 112, 450, 112, .TextBox1, 1

    End Dialog

    Dim dlg As UserDialog
    dlg.TextBox1 = ""



    If(QueryDadosRotinaCartao.FieldByName("ROTINACARTAO_CODIGO").AsString <>"")Then
    dlg.TextBox1 = dlg.TextBox1 + "Código da Rotina: " + QueryDadosRotinaCartao.FieldByName("ROTINACARTAO_CODIGO").AsString + Chr(13) + Chr(10)
  Else
    dlg.TextBox1 = dlg.TextBox1 + "Código da Rotina:- " + Chr(13) + Chr(10)
  End If

  If(QueryDadosRotinaCartao.FieldByName("ROTINACARTAO_DESCRICAO").AsString <>"")Then
  dlg.TextBox1 = dlg.TextBox1 + "Descrição: " + QueryDadosRotinaCartao.FieldByName("ROTINACARTAO_DESCRICAO").AsString + Chr(13) + Chr(10)
Else
  dlg.TextBox1 = "Descricao:-" + Chr(13) + Chr(10)
End If

If(Not QueryDadosRotinaCartao.FieldByName("ROTINACARTAO_DATAROTINA").IsNull)Then
dlg.TextBox1 = dlg.TextBox1 + "Data Rotina: " + Str(Format(QueryDadosRotinaCartao.FieldByName("ROTINACARTAO_DATAROTINA").AsDateTime, "dd/mm/yyyy")) + Chr(13) + Chr(10)
Else
  dlg.TextBox1 = "Data Rotina:-" + Chr(13) + Chr(10)
End If


On Error GoTo cancel
Dialog dlg

'Query que conta as quantidades

'		QueryContaQtdComunicado.Add("SELECT	RCC.TIPOARQUIVO,COUNT(*) QTD							")
'		QueryContaQtdComunicado.Add("FROM SAM_ROTINACARTAO 				RC			")
'		QueryContaQtdComunicado.Add("	JOIN SAM_ROTINACARTAO_CARTAO		RCC	ON(RCC.ROTINACARTAO = RC.HANDLE)")
'		QueryContaQtdComunicado.Add("	JOIN SAM_BENEFICIARIO_CARTAOIDENTIF	BCI	ON(BCI.HANDLE = RCC.CARTAOIDENTIFICACAO)")
'		QueryContaQtdComunicado.Add("	JOIN SAM_BENEFICIARIO 			B	ON(B.HANDLE = BCI.BENEFICIARIO)")
'		QueryContaQtdComunicado.Add("	JOIN SAM_FAMILIA	 			F	ON(F.HANDLE = B.FAMILIA)")
'		QueryContaQtdComunicado.Add("WHERE (RC.HANDLE = :P_ROTINACARTAO_HANDLE)")
'		QueryContaQtdComunicado.Add("AND (BCI.SITUACAO<>'C') AND (BCI.DATAEMISSAO IS NOT NULL)")

QueryContaQtdComunicado.Add("SELECT	RCC.TIPOARQUIVO,COUNT(*) QTD")
QueryContaQtdComunicado.Add("  FROM SAM_ROTINACARTAO RC,")
QueryContaQtdComunicado.Add("       SAM_ROTINACARTAO_CARTAO RCC,")
QueryContaQtdComunicado.Add("       SAM_BENEFICIARIO_CARTAOIDENTIF	BCI,")
QueryContaQtdComunicado.Add("       SAM_BENEFICIARIO B,")
QueryContaQtdComunicado.Add("       SAM_FAMILIA F")
QueryContaQtdComunicado.Add(" WHERE (RC.HANDLE = :P_ROTINACARTAO_HANDLE)")
QueryContaQtdComunicado.Add("   AND (BCI.SITUACAO<>'C') AND (BCI.DATAEMISSAO IS NOT NULL)")
QueryContaQtdComunicado.Add("   AND (RCC.ROTINACARTAO = RC.HANDLE)")
QueryContaQtdComunicado.Add("   AND (BCI.HANDLE = RCC.CARTAOIDENTIFICACAO)")
QueryContaQtdComunicado.Add("   AND (B.HANDLE = BCI.BENEFICIARIO)")
QueryContaQtdComunicado.Add("   AND (F.HANDLE = B.FAMILIA)")

Select Case dlg.Group1
  Case 0 'COMUNICADO DE INGRESSO
    QueryContaQtdComunicado.Add("AND (RCC.TIPOARQUIVO='2' OR RCC.TIPOARQUIVO='3') ")' TIPOARQUIVO=BENEFICIÁRIO OU FAMÍLIA")

End Select

QueryContaQtdComunicado.Add("GROUP BY RCC.TIPOARQUIVO")
QueryContaQtdComunicado.ParamByName("P_ROTINACARTAO_HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
QueryContaQtdComunicado.Active = False
QueryContaQtdComunicado.Active = True

'Acumulador de quantidade de cartões

TextoFinal = ""
TextoFinal = "Resumo da Geração :" + Chr(13) + Chr(10)
TextoFinal = TextoFinal + " " + Chr(13) + Chr(10)


QtdCartao = 0

QueryContaQtdComunicado.First
dlg.textBox1 = dlg.textBox1 + Chr(13) + Chr(10)
While(Not(QueryContaQtdComunicado.EOF))
QtdCartao = QtdCartao + QueryContaQtdComunicado.FieldByName("QTD").AsInteger
Select Case QueryContaQtdComunicado.FieldByName("TIPOARQUIVO").AsString
  Case "1" 'contrato
    TextoFinal = TextoFinal + "          Para Contrato:" + Str(QueryContaQtdComunicado.FieldByName("QTD").AsInteger) + Chr(13) + Chr(10)

  Case "2" 'família
    TextoFinal = TextoFinal + "          Para Família:" + Str(QueryContaQtdComunicado.FieldByName("QTD").AsInteger) + Chr(13) + Chr(10)

  Case "3" 'beneficiário
    TextoFinal = TextoFinal + "          Para Beneficiário:" + Str(QueryContaQtdComunicado.FieldByName("QTD").AsInteger) + Chr(13) + Chr(10)

End Select
QueryContaQtdComunicado.Next
Wend

dlg.TextBox1 = dlg.textbox1 + "Total de Comunicados:" + Str(QtdCartao) + Chr(13) + Chr(10)

'Seleciona qual relatório foi selecionado no DLG

Select Case dlg.Group1
  Case 0 'COMUNICADO DE INGRESSO
    QueryBuscaRelatorioComunicadoHandle.ParamByName("P_CODIGO").Value = "BEN006"
    FormFiltroTexto = "Comunicado de Ingresso"

  Case 1 'COMUNICADO DE RENOVAÇÃO
    QueryBuscaRelatorioComunicadoHandle.ParamByName("P_CODIGO").Value = "BEN005"
    FormFiltroTexto = "Comunicado de Renovação"

End Select


QueryBuscaRelatorioComunicadoHandle.Active = False
QueryBuscaRelatorioComunicadoHandle.Active = True
RelatorioComunicadoHandle = QueryBuscaRelatorioComunicadoHandle.FieldByName("HANDLE").AsInteger

'gravar os resultados do comunicado para o log







'pegar o nome do arquivo para o qual será gerado o comunicado
Dim FiltroArquivo As String
Dim Interface As Object
Dim HandleFiltro As Long
Dim PosInicial As Long
Dim PosFinal As Long
Dim vExtensao As String
Dim SQLFiltro As Object
Dim QAtualiza As Object
Set QAtualiza = NewQuery



Set Interface = CreateBennerObject("SamFiltro.Filtro")
VOLTA :
HandleFiltro = Interface.Exec(CurrentSystem, CurrentUser, 16, "ARQUIVO", FormFiltroTexto)
Set Interface = Nothing


If(HandleFiltro >0)Then
'Busca pelo RFFiltro selecionado
Set SQLFiltro = NewQuery
SQLFiltro.Add("SELECT ARQUIVO FROM RF_FILTRO WHERE HANDLE=" + Str(HandleFiltro))
SQLFiltro.Active = True

PosInicial = InStr(SQLFiltro.FieldByName("ARQUIVO").AsString, ".") + 1
PosFinal = (Len(SQLFiltro.FieldByName("ARQUIVO").AsString) - PosInicial) + 1

vExtensao = UCase((Mid(SQLFiltro.FieldByName("ARQUIVO").AsString, PosInicial, PosFinal)))
If vExtensao <>"DAT" Then
  bsShowMessage("A extensão do arquivo deve ser *.DAT", "I")
  GoTo VOLTA
End If

'Atribui valores às variaveis de filtro
FiltroArquivo = SQLFiltro.FieldByName("ARQUIVO").AsString

Set SQLFiltro = Nothing


On Error GoTo LabelTeste
ReportExport(RelatorioComunicadoHandle, "", FiltroArquivo, False, False)
'	ReportPreview(RelatorioComunicadoHandle,"",False,False)
LabelTeste :

If(UserParam <>QtdCartao)Then
GoTo LabelErro

Else
  'MsgBox FormFiltroTexto+" Gerado com sucesso:"+UserParam+" Registro(s) do Total de "+Str(QtdCartao)+" Registro(s)."+Chr(13)+"Arquivo:"+FiltroArquivo+Chr(13)+"Resumo da Geração:"+Chr(13)+TextoFinal
  Begin Dialog UserDialog 470, 231 ' %GRID:10,7,1,1
    OKButton 190, 200, 90, 21
    Text 20, 8, 600, 14, FormFiltroTexto + " Gerado com sucesso!"
    Text 30, 26, 600, 14, "Gerado(s) " + UserParam + " Registro(s) do total de " + Str(QtdCartao) + " Registro(s).", .Text2
    Text 30, 44, 600, 14, "Arquivo:" + FiltroArquivo, .Text3
    TextBox 20, 67, 430, 112, .TextBox1, 1

    End Dialog

    vMessage = vMessage + FormFiltroTexto + " Gerado com sucesso! " + Chr(13) + "Gerado(s) " + UserParam + " Registro(s) do total de " + Str(QtdCartao) _
                + " Registro(s)." + Chr(13) + "Arquivo:" + FiltroArquivo + Chr(13)




    Dim dlg2 As UserDialog
    dlg2.TextBox1 = TextoFinal

    Dialog dlg2

    QAtualiza.Clear
    QAtualiza.Active = False
    QAtualiza.Add("UPDATE SAM_ROTINACARTAO SET OCORRENCIAS =:OCORRENCIAS WHERE HANDLE =:HANDLE")
    QAtualiza.ParamByName("OCORRENCIAS").AsMemo = QueryDadosRotinaCartao.FieldByName("OCORRENCIAS").Value + " " + vMessage + " " + TextoFinal
    QAtualiza.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    QAtualiza.ExecSQL


    GoTo LabelFim

  End If

LabelErro :
  If(UserParam >QtdCartao)Then
  bsShowMessage("Erro Geração " + FormFiltroTexto + " . Gerados " + Str(0) + " Registro(s) do Total de " + Str(QtdCartao) + " registros.", "E")
  vMessage = "Erro Geração " + FormFiltroTexto + " . Gerados " + Str(0) + " Registro(s) do Total de " + Str(QtdCartao) + " registros."
  ElseIf(UserParam <QtdCartao)Then
  bsShowMessage("Erro Geração " + FormFiltroTexto + " . Gerados Parcialmente " + Str(UserParam) + " Registro(s) do Total de " + Str(QtdCartao) + " registros.", "E")
  vMessage = "Erro Geração " + FormFiltroTexto + " . Gerados Parcialmente " + Str(UserParam) + " Registro(s) do Total de " + Str(QtdCartao) + " registros."
End If

QAtualiza.Clear
QAtualiza.Active = False
QAtualiza.Add("UPDATE SAM_ROTINACARTAO SET OCORRENCIAS =:OCORRENCIAS WHERE HANDLE =:HANDLE")
QAtualiza.ParamByName("OCORRENCIAS").AsMemo = QueryDadosRotinaCartao.FieldByName("OCORRENCIAS").Value + " " + vMessage + " " + TextoFinal
QAtualiza.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
QAtualiza.ExecSQL

Set QAtualiza = Nothing


LabelFim :
End If
Set QueryDadosRotinaCartao = Nothing
Set QueryBuscaRelatorioComunicadoHandle = Nothing
Set QueryContaQtdComunicado = Nothing
Set QAtualiza = Nothing

cancel :
Set QueryDadosRotinaCartao = Nothing
Set QueryBuscaRelatorioComunicadoHandle = Nothing
Set QueryContaQtdComunicado = Nothing
'fim




End Sub

Public Sub BOTAODESBLOQUEAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
    bsShowMessage("A Geração ainda não foi processada", "I")
    Exit Sub
  End If
  Dim S As Object

  Set S = NewQuery

  '  S.Add("SELECT COUNT(R.HANDLE) PARAMETROS")
  '  S.Add("FROM SAM_ROTINACARTAO R")
  '  S.Add("     JOIN sam_rotinacartao_cartao C ON")
  '  S.Add("     (C.ROTINACARTAO = R.handle)")
  '  S.Add("     JOIN sam_beneficiario_cartaoidentif D ON")
  '  S.Add("     (D.handle = c.CARTAOIDENTIFICACAO)")
  '  S.Add("WHERE R.handle = :HROTINACARTAO")
  '  S.Add("  AND D.SITUACAO = 'B'               ")
  S.Add("SELECT COUNT(R.HANDLE) PARAMETROS")
  S.Add("FROM SAM_ROTINACARTAO R,")
  S.Add("     SAM_ROTINACARTAO_cartao C,")
  S.Add("     SAM_BENEFICIARIO_CARTAOIDENTIF D")
  S.Add("WHERE R.handle = :HROTINACARTAO")
  S.Add("  AND (C.ROTINACARTAO = R.handle)")
  S.Add("  AND (D.handle = c.CARTAOIDENTIFICACAO)")
  S.Add("  AND D.SITUACAO = 'B'               ")


  S.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  S.Active = True

  If S.FieldByName("PARAMETROS").Value = 0 Then
    bsShowMessage("Não há Cartões para Desbloquear!!.", "I")
    Exit Sub
  End If

  '  If MsgBox("Confirma o desbloqueio dos cartões ?",vbYesNo,"Rotina de Cartões")=vbYes Then
 ' Set Obj = CreateBennerObject("SAMROTINACARTAO.Geral")
 ' Obj.Desbloquear(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 0)
  Set Obj = Nothing
  '  End If

  If VisibleMode Then

    Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")
    Obj.Desbloquear(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Else

    Dim vsMensagemErro As String
    Dim viRetorno As Long


    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_Desbloquear", _
                                     "Rotina de Desbloqueio de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO", _
                                     "SITUACAODESBLOQUEIO", _
                                     "", _
                                     "", _
                                     "P", _
                                     True, _
                                      vsMensagemErro, _
                                     Null)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

  End If


  WriteAudit("D", HandleOfTable("SAM_ROTINACARTAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Cartões - Desbloqueio")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOEXPORTAR_OnClick()

  'SMS 58772 - inicio
  UserVar("Sender") = "BOTAOEXPORTAR"
  UserParam         = CurrentQuery.FieldByName("HANDLE").AsInteger
  'SMS 58772 - fim
  Dim Obj As Object
  Dim ParametrosBenef As BPesquisa

  Set ParametrosBenef = NewQuery

  ParametrosBenef.Active = False
  ParametrosBenef.Add(" SELECT * 						  ")
  ParametrosBenef.Add("   FROM SAM_PARAMETROSBENEFICIARIO ")
  ParametrosBenef.Active = True

  If VisibleMode Then

    Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")
    Obj.Exportar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	If (ParametrosBenef.FieldByName("DESBLOQUEARAPOSEXPORTACAO").AsString = "S") Then
      Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")
      Obj.Desbloquear(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
	End If

  Else

    Dim vsMensagemErro As String
    Dim viRetorno As Long
	Dim viRetornoDesbloqueio As Long

	If(CurrentQuery.FieldByName("ARQUIVOEXP").AsString = "")Then
		bsShowMessage("Falta preencher o campo 'Nome do Arquivo para exportação' com o caminho para salvar o documento.", "E")
	End If

	SessionVar("DIRETORIOEXPORTACAOCARTAO") = CurrentQuery.FieldByName("ARQUIVOEXP").AsString

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_Exportar", _
                                     "Rotina de Exportação de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO", _
                                     "SITUACAOEXPORTACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     True, _
                                      vsMensagemErro, _
                                     Null)

	If (ParametrosBenef.FieldByName("DESBLOQUEARAPOSEXPORTACAO").AsString = "S") Then
		viRetornoDesbloqueio = Obj.ExecucaoImediata(CurrentSystem, _
													"BSBEN009", _
													"RotinaCartao_Desbloquear", _
													"Rotina de Desbloqueio de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
													CurrentQuery.FieldByName("HANDLE").AsInteger, _
													"SAM_ROTINACARTAO", _
													"SITUACAODESBLOQUEIO", _
													"", _
													"", _
													"P", _
													True, _
													vsMensagemErro, _
													Null)
	End If

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

  End If

  Set ParametrosBenef = Nothing
  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINACARTAO")

End Sub

Public Sub BOTAOFATURA_OnClick()

  Dim Interface As Object
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim S As Object

  Set S = NewQuery

  '  S.Add("SELECT COUNT(R.HANDLE) PARAMETROS")
  '  S.Add("FROM SAM_ROTINACARTAO_CARTAO R")
  '  S.Add("     JOIN SAM_ROTINACARTAO_FATURA C ON")
  '  S.Add("     (    C.ROTINACARTAO = R.ROTINACARTAO)")
  '  S.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
  '  S.Add("  AND C.SITUACAO = 'A'               ")
  S.Add("SELECT COUNT(R.HANDLE) PARAMETROS")
  S.Add("FROM SAM_ROTINACARTAO_CARTAO R, ")
  S.Add("     SAM_ROTINACARTAO_FATURA C")
  S.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
  S.Add("  AND (C.ROTINACARTAO = R.ROTINACARTAO)")
  S.Add("  AND C.SITUACAO = 'A'               ")

  S.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  S.Active = True

  If S.FieldByName("PARAMETROS").Value = 0 Then
    bsShowMessage("Não há Parâmetros para Faturamento em Aberto.", "I")
    Exit Sub
  End If

  Dim SQL As Object
  Dim TextoQtdCartoes As String

  Set SQL = NewQuery

  'Cartões do arquvio por contrato
  'SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
  'SQL.Add("FROM SAM_ROTINACARTAO_CARTAO R")
  'SQL.Add("     JOIN SAM_BENEFICIARIO_CARTAOIDENTIF C ON")
  'SQL.Add("           (C.HANDLE = R.CARTAOIDENTIFICACAO)")
  'SQL.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
  'SQL.Add("      AND C.SITUACAO = 'N'           ")
  'SQL.Add("  AND R.TIPOARQUIVO = '1'")
  SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
  SQL.Add("FROM SAM_ROTINACARTAO_CARTAO R, ")
  SQL.Add("     SAM_BENEFICIARIO_CARTAOIDENTIF C")
  SQL.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
  SQL.Add("  AND (C.HANDLE = R.CARTAOIDENTIFICACAO)")
  SQL.Add("  AND C.SITUACAO = 'N'           ")
  'SQL.Add("  AND R.TIPOARQUIVO = '1'")

  SQL.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If SQL.FieldByName("CARTOES").Value = 0 Then
    bsShowMessage("Não há Cartões a serem Faturados.", "I")
    Exit Sub
  End If

  If SQL.EOF Then
    bsShowMessage("Não há Cartões Gerados.", "I")
    Exit Sub
  End If


  If VisibleMode Then

        Set Interface = CreateBennerObject("SAMROTINACARTAO.Geral")
        Interface.Fatura(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
        Set Interface = Nothing

        WriteAudit("F", HandleOfTable("SAM_ROTINACARTAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Cartões - Faturamento")

  Else
         Dim vsMensagemErro As String
   		 Dim viRetorno As Long


         Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	 viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_GerarFatura", _
                                     "Rotina de Fatura de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO", _
                                     "SITUACAOFATURA", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                      vsMensagemErro, _
                                     Null)


        If viRetorno = 0 Then
  			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
    		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      	End If
  End If


  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub


Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("tabtipogeracao").AsInteger = 2 Then
    If(Not CurrentQuery.FieldByName("USUARIOEXPORTACAO").IsNull) _
       Or(Not CurrentQuery.FieldByName("USUARIODESBLOQUEIO").IsNull) _
       Or(Not CurrentQuery.FieldByName("USUARIOFATURAR").IsNull)Then
        bsShowMessage("Não é mais permitido gerar cartão avulso para essa rotina.", "I")
    Exit Sub
    End If
  Else
    If Not CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
      bsShowMessage("A Geração já foi processada", "I")
      Exit Sub
    End If
  End If

If CurrentQuery.FieldByName("tabtipogeracao").AsInteger = 1 Then
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT COUNT(HANDLE) QTDCONTRATOS ")
  SQL.Add("FROM SAM_ROTINACARTAO_CONTRATO")
  SQL.Add("WHERE ROTINACARTAO = :HROTINACARTAO")
  SQL.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True '
  If SQL.FieldByName("QTDCONTRATOS").AsInteger = 0 Then
    bsShowMessage("Rotina sem Contrato(s) selecionado(s).", "I")
    Exit Sub
  End If
End If

Set SQL = Nothing

If CurrentQuery.FieldByName("tabtipogeracao").AsInteger = 4 Then
  Dim SQL2 As Object
  Set SQL2 = NewQuery
  SQL2.Add("SELECT COUNT(HANDLE) QTDCONTRATOS ")
  SQL2.Add("  FROM CA_ROTSOLICIT")
  SQL2.Add(" WHERE ROTINACARTAO = :HROTINACARTAO")
  SQL2.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL2.Active = True '
  If SQL2.FieldByName("QTDCONTRATOS").AsInteger = 0 Then
    bsShowMessage("Rotina sem solicitações geradas.", "I")
    Exit Sub
  End If
End If

'Set Obj = CreateBennerObject("SAMROTINACARTAO.Geral")
'Obj.Gerar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

          If VisibleMode Then

		    Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")
		    Obj.Gerar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

		  Else

		    Dim vsMensagemErro As String
   			Dim viRetorno As Long

		    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    		viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_Gerar", _
                                     "Rotina de Geração de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                      vsMensagemErro, _
                                     Null)


	    	If viRetorno = 0 Then
      				bsShowMessage("Processo enviado para execução no servidor!", "I")
		  	Else
      				bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	  	End If

          End If






Set Obj = Nothing

WriteAudit("G", HandleOfTable("SAM_ROTINACARTAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Cartões - Geração")

CurrentQuery.Active = False
CurrentQuery.Active = True
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
Dim RelatorioHandle As Integer

  UserVar("Sender") = "BOTAOIMPRIMIR"

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("usuariogeracao").IsNull Then
    bsShowMessage("Rotina não Gerada", "I")
    Exit Sub
  End If

  If(CurrentQuery.FieldByName("TABTIPOGERACAO").AsInteger = 2)And(CurrentQuery.FieldByName("USUARIODESBLOQUEIO").IsNull)Then
    bsShowMessage("É necessário desbloquear cartão!", "I")
    Exit Sub
  End If

  Dim qVerificaTipoCartao As Object
  Set qVerificaTipoCartao = NewQuery

  qVerificaTipoCartao.Active = False
  qVerificaTipoCartao.Add("SELECT A.HANDLE, A.BENEFICIARIO			 ")
  qVerificaTipoCartao.Add("  FROM SAM_BENEFICIARIO_CARTAOIDENTIF A,  ")
  qVerificaTipoCartao.Add("       SAM_ROTINACARTAO_CARTAO B          ")
  qVerificaTipoCartao.Add(" WHERE A.HANDLE = B.CARTAOIDENTIFICACAO   ")
  qVerificaTipoCartao.Add("   AND B.ROTINACARTAO = :HROTINA          ")
  qVerificaTipoCartao.Add("   AND A.TIPOCARTAO IS NOT NULL           ")
  qVerificaTipoCartao.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaTipoCartao.Active = True

  'sms 22944 - fernando
  If Not qVerificaTipoCartao.FieldByName("HANDLE").IsNull Then
    Dim Interface As Object
    Dim SQLFiltro As Object
    Dim QueryBuscaHRelatorio As Object
    Dim HandleFiltro As Integer

    Set Interface =CreateBennerObject("SamFiltro.Filtro")

    HandleFiltro =Interface.Exec(CurrentSystem,CurrentUser,1,"TIPOCARTAO.nl.ob|TIPODEPENDENTE","Escolha o tipo de cartão desejado para ser impresso")

    Set SQLFiltro =NewQuery
    SQLFiltro.Add("SELECT A.TIPOCARTAO, B.CAMPO                                ")
    SQLFiltro.Add("  FROM RF_FILTRO A					   					   ")
    SQLFiltro.Add("  LEFT JOIN RF_FILTRO_SELECAO B ON (A.HANDLE = B.RF_FILTRO) ")
    SQLFiltro.Add(" WHERE A.HANDLE = " +Str(HandleFiltro))
    SQLFiltro.Active =True

    Set QueryBuscaHRelatorio = NewQuery

    QueryBuscaHRelatorio.Active = False
    QueryBuscaHRelatorio.Add("SELECT RELATORIOESPECIFICO  ")
    QueryBuscaHRelatorio.Add("  FROM SAM_TIPOCARTAO       ")
    QueryBuscaHRelatorio.Add(" WHERE HANDLE = :TIPOCARTAO ")
    QueryBuscaHRelatorio.ParamByName("TIPOCARTAO").AsInteger = SQLFiltro.FieldByName("TIPOCARTAO").AsInteger

    QueryBuscaHRelatorio.Active = True

    If QueryBuscaHRelatorio.FieldByName("RELATORIOESPECIFICO").IsNull Then
      bsShowMessage("Falta configurar o relatório de impressão no tipo de cartão do beneficiário!","I")
      Exit Sub
    End If

    RelatorioHandle = QueryBuscaHRelatorio.FieldByName("RELATORIOESPECIFICO").AsInteger

    SessionVar("HANDLEROTINACARTAO") = Str(CurrentQuery.FieldByName("HANDLE").AsInteger)
    UserVar("HANDLEFILTRO")          = Str(HandleFiltro)
    UserVar("HBENEF")                = ""

	Dim samBeneficiarioBLL As CSBusinessComponent
	Set samBeneficiarioBLL = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.SamBeneficiarioCartaoIdentifBLL, Benner.Saude.Beneficiarios.Business")


    If InStr("CACHE",SQLServer) > 0 Then

       Dim Interface2 As Object
       Set Interface2 =CreateBennerObject("SAMROTINACARTAO.Beneficiario")
       Interface2.ImprimeRelatorioCartao(CurrentSystem)

    Else
	   UserVar("HANDLEBENEF") = qVerificaTipoCartao.FieldByName("BENEFICIARIO").AsString
       ReportPreview(RelatorioHandle, "", False, False)

    End If

    Set samBeneficiarioBLL = Nothing
    Set SQLFiltro =Nothing
    Set Interface =Nothing

  Else
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT RELATORIOCARTAO,TABTIPOLEIAUTE FROM SAM_PARAMETROSBENEFICIARIO")
    SQL.Active = True

    If SQL.FieldByName("TABTIPOLEIAUTE").AsInteger <> 4 Then
      If Not SQL.FieldByName("RELATORIOCARTAO").IsNull Then
        UserParam = CurrentQuery.FieldByName("HANDLE").AsInteger
        ReportPreview(SQL.FieldByName("RELATORIOCARTAO").AsInteger, "", False, False)
        Set SQL = Nothing
      Else
        bsShowMessage("Defina relatório de impressão ou Mude o tipo de leiaute para variável!", "I")
      End If
    Else
      Dim vLogRetorno As String
      Dim Obj As Object
      Set Obj = CreateBennerObject("SAMIMPRIMIRCARTAO.Exportar")
      Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 2, vLogRetorno)
      Set Obj = Nothing
    End If
  End If

  Set qVerificaTipoCartao = Nothing
End Sub


Public Sub BOTAOPROTOCOLO_OnClick()

  '++++++++teste mauricio ++++++++
  '  Set Obj=CreateBennerObject("SAMROTINACARTAO.Geral")
  '  Obj.Gerar(CurrentQuery.FieldByName("HANDLE").AsInteger)
  '  Set Obj=Nothing

  '  Exit Sub
  Dim RelatorioProtocoloHandle As Long
  Dim RotinaCartaoAtualHandle As Long 'Handle da Rotina Carta Atual
  Dim QueryBuscaHandleProtocolo As Object 'busca handle do relatorio protocolo de entrega de cartões

  Set QueryBuscaHandleProtocolo = NewQuery

  QueryBuscaHandleProtocolo.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'BEN043'")
  'QueryBuscaHandleProtocolo.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'SFN-DC-999'")

  QueryBuscaHandleProtocolo.Active = False
  QueryBuscaHandleProtocolo.Active = True

  RelatorioProtocoloHandle = QueryBuscaHandleProtocolo.FieldByName("HANDLE").AsInteger


  Set QueryBuscaHandleProtocolo = Nothing

  'Parâmetro global
  UserVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString

  'no "WHERE" da exportação seleciona-se os beneficiarios que estão abaixo da carga de rotina
  ReportPreview(RelatorioProtocoloHandle, "", False, False)
  'ReportPrint(RelatorioProtocoloHandle,"",False,False)

End Sub

Public Sub CONTRATOFINAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraContrato(CurrentQuery.FieldByName("CONTRATOFINAL").Value)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOFINAL").Value = vHandle
  End If
End Sub

Public Sub CONTRATOINICIAL_OnChange()
  CurrentQuery.FieldByName("CONTRATOFINAL").Value = CurrentQuery.FieldByName("CONTRATOINICIAL").AsInteger
End Sub

Public Sub CONTRATOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraContrato(CurrentQuery.FieldByName("CONTRATOINICIAL").Value)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOINICIAL").Value = vHandle
    CONTRATOINICIAL_OnChange
  End If

End Sub

Public Sub BOTAORETORNO_OnClick()
  Dim Obj As Object
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Set Obj = CreateBennerObject("SAMROTINACARTAO.Geral")
  Obj.Retorno(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing

  WriteAudit("R", HandleOfTable("SAM_ROTINACARTAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Cartões - Retorno")

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub TABLE_AfterScroll()
  'Pinheiro - Alteração acrescentadoa porque o campo situação não gravava o valor padrão

'SMS 52120 - Marcelo Barbosa - 26/12/2005
'Verifica se a carga correspondente, através do codigo interno(definido no Builder), é do Modelo de agendamento
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina

  If (VisibleMode And NodeInternalCode <> 500) Or (WebMode And WebMenuCode <> "T5140") Then
    If Not CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
		TABTIPOGERACAO.ReadOnly = True
	Else
		TABTIPOGERACAO.ReadOnly = False
    End If
  End If

If VisibleMode Then
	If (VisibleMode And NodeInternalCode = 500) Then

  	TABTIPOROTINA.Pages(0).Visible = False
  	TABTIPOROTINA.Pages(1).Visible = True
  	'CurrentQuery.FieldByName("TABTIPOROTINA").Value = 2
  	BOTAOAGENDAR.Visible = False
  	BOTAOCANCELAFATURAMENTO.Visible = False
  	BOTAOCANCELAR.Visible = False
  	BOTAOCOMUNICADO.Visible = False
  	BOTAODESBLOQUEAR.Visible = False
  	BOTAOEXPORTAR.Visible = False
  	BOTAOFATURA.Visible = False
  	BOTAOGERAR.Visible = False
  	BOTAOIMPRIMIR.Visible = False
  	BOTAOPROTOCOLO.Visible = False

  	REFERENCIA.ReadOnly = False
  	TIPOCARTAO.ReadOnly = False
  	DESCRICAO.ReadOnly = False

	Else

	  TABTIPOROTINA.Pages(0).Visible = True
  	  TABTIPOROTINA.Pages(1).Visible = False
  'CurrentQuery.FieldByName("TABTIPOROTINA").Value = 1

  If CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
    REFERENCIA.ReadOnly = False
    TIPOCARTAO.ReadOnly = False
    DESCRICAO.ReadOnly = False
  Else
    REFERENCIA.ReadOnly = True
    TIPOCARTAO.ReadOnly = True
    DESCRICAO.ReadOnly = True
  End If

  If (CurrentQuery.FieldByName("USUARIOEXPORTACAO").IsNull) Then
    TABEXPORTACAO.ReadOnly = False
    ARQUIVOCONTRATO.ReadOnly = False
    ARQUIVOFAMILIA.ReadOnly = False
    ARQUIVOBENEFICIARIO.ReadOnly = False
    ARQUIVO.ReadOnly = False
  Else
    TABEXPORTACAO.ReadOnly = True
    ARQUIVOCONTRATO.ReadOnly = True
    ARQUIVOFAMILIA.ReadOnly = True
    ARQUIVOBENEFICIARIO.ReadOnly = True
    ARQUIVO.ReadOnly = True
  End If


  ROTULOGERACAO.Text = ""
  If Not CurrentQuery.FieldByName("USUARIOEXPORTACAO").IsNull Then
    Dim SQL As Object
    Dim TextoQtdCartoes As String

    Set SQL = NewQuery


   'SMS 51641
    SQL.Clear
    'SMS 84056 - Débora Rebello (CASSI) - 03/07/2007 - inicio

    'If InStr(SQLServer, "ORACLE")>0 Then
    '   SQL.Add("SELECT /*+rule*/ COUNT(R.HANDLE) CARTOES")
    'Else
      'SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
    'End If
    'fim sms 51641

    SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
    'SMS 84056 - Débora Rebello (CASSI) - 03/07/2007 - fim

    SQL.Add("FROM SAM_ROTINACARTAO_CARTAO R, ")
    SQL.Add("     SAM_BENEFICIARIO_CARTAOIDENTIF C")
    SQL.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
    SQL.Add("  AND (C.HANDLE = R.CARTAOIDENTIFICACAO)")
    SQL.Add("  AND R.TIPOARQUIVO = '3'            ")
    SQL.Add("  AND C.DATAEMISSAO IS NOT NULL      ")

    SQL.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If SQL.FieldByName("CARTOES").AsInteger >0 Then
      TextoQtdCartoes = " - Contrato: " + SQL.FieldByName("CARTOES").AsString
    End If

  'SMS 51641
   SQL.Clear
   'SMS 84056 - Débora Rebello (CASSI) - 03/07/2007 - inici0

   'If InStr(SQLServer, "ORACLE")>0 Then
   '  SQL.Add("SELECT /*+rule*/ COUNT(R.HANDLE) CARTOES")
   'Else
   ' SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
   'End If
   'fim sms 51641
   SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
   'SMS 84056 - Débora Rebello (CASSI) - 03/07/2007 - fim

   SQL.Add("FROM SAM_ROTINACARTAO_CARTAO R, ")
   SQL.Add("     SAM_BENEFICIARIO_CARTAOIDENTIF C")
   SQL.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
   SQL.Add("  AND (C.HANDLE = R.CARTAOIDENTIFICACAO)")
   SQL.Add("  AND R.TIPOARQUIVO = '1'            ")
   SQL.Add("  AND C.DATAEMISSAO IS NOT NULL      ")

    SQL.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If SQL.FieldByName("CARTOES").AsInteger >0 Then
      TextoQtdCartoes = TextoQtdCartoes + " - Família: " + SQL.FieldByName("CARTOES").AsString
    End If

    'Cartões do arquvio por beneficiário
    'SQL.Clear
    'SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
    'SQL.Add("FROM SAM_ROTINACARTAO_CARTAO R")
    'SQL.Add("     JOIN SAM_BENEFICIARIO_CARTAOIDENTIF C ON")
    'SQL.Add("     (    C.HANDLE = R.CARTAOIDENTIFICACAO)")
    'SQL.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
    'SQL.Add("  AND R.TIPOARQUIVO = '3'            ")
    'SQL.Add("  AND C.DATAEMISSAO IS NOT NULL      ")

   'SMS 51641
    SQL.Clear
    'SMS 84056 - Débora Rebello (CASSI) - 03/07/2007 - inicio

    'If InStr(SQLServer, "ORACLE")>0 Then
    '  SQL.Add("SELECT /*+rule*/ COUNT(R.HANDLE) CARTOES")
    'Else
    '  SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
    'End If
   'fim sms 51641

    SQL.Add("SELECT COUNT(R.HANDLE) CARTOES")
   'SMS 84056 - Débora Rebello (CASSI) - 03/07/2007 - fim
    SQL.Add("FROM SAM_ROTINACARTAO_CARTAO R, ")
    SQL.Add("     SAM_BENEFICIARIO_CARTAOIDENTIF C")
    SQL.Add("WHERE R.ROTINACARTAO = :HROTINACARTAO")
    SQL.Add("  AND (C.HANDLE = R.CARTAOIDENTIFICACAO)")
    SQL.Add("  AND R.TIPOARQUIVO = '2'            ")
    SQL.Add("  AND C.DATAEMISSAO IS NOT NULL      ")
    SQL.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If SQL.FieldByName("CARTOES").AsInteger >0 Then
      TextoQtdCartoes = TextoQtdCartoes + " - Beneficiário: " + SQL.FieldByName("CARTOES").AsString
    End If

    If TextoQtdCartoes <>"" Then
      ROTULOGERACAO.Text = "CARTÕES EXPORTADOS" + TextoQtdCartoes
    End If

    SQL.Active = False
    Set SQL = Nothing
  End If
End If
End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If (VisibleMode And NodeInternalCode <> 500) Or (WebMode And WebMenuCode <> "T5140") Then
    If Not CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
      bsShowMessage("A Geração já foi processada. Não é possível excluir o registro", "E")
      CanContinue = False
    End If
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Q As BPesquisa

  If (VisibleMode And NodeInternalCode <> 500) Or (WebMode And WebMenuCode <> "T5140") Then

    If CurrentQuery.FieldByName("TABTIPOGERACAO").AsInteger = 2 Then
      CurrentQuery.FieldByName("REFERENCIA").Clear
      CurrentQuery.FieldByName("TIPOCARTAO").Clear
    End If

    If(CurrentQuery.FieldByName("rotsolicitparam").IsNull)And _
      (BOTAOGERAR.Visible = True)Then
      If CurrentQuery.FieldByName("tabtipogeracao").Value = 4 Then
        bsShowmessage("Tipo de geração 'Solicitação' é gerado automaticamente pela Central de Atendimento", "E")
        CanContinue = False
        Exit Sub
      End If
    End If

    If Not(CurrentQuery.FieldByName("rotsolicitparam").IsNull)Then
      Dim SQL As Object
      Set SQL = NewQuery
      SQL.Add("SELECT B.SITUACAO ")
      SQL.Add("FROM CA_ROTSOLICITPARAM A,")
      SQL.Add("     CA_ROTSOLICIT B ")
      SQL.Add("WHERE A.HANDLE = :HROTPARAM AND B.HANDLE = A.ROTSOLICIT")
      SQL.ParamByName("HROTPARAM").Value = CurrentQuery.FieldByName("rotsolicitparam").AsInteger
      SQL.Active = True '

  'If SQL.FieldByName("SITUACAO").AsString <>"A" Then
  '  MsgBox("Situação da rotina de solicitação não permitida para alteração.")
  '  CanContinue =False
  '  RefreshNodesWithTable("SAM_ROTINACARTAO")
  '
  '  Exit Sub
  'End If

     If CurrentQuery.FieldByName("tabtipogeracao").Value <>4 Then
       bsShowMessage("Tipo de geração 'Solicitação' é gerado automaticamente pela Central de Atendimento", "E")
       CanContinue = False
       Exit Sub
     End If
   End If

End If

Set Q = NewQuery
Q.Add("SELECT TABTIPOLEIAUTE FROM SAM_PARAMETROSBENEFICIARIO")
Q.Active = True

If Q.FieldByName("TABTIPOLEIAUTE").AsInteger = 4 And CurrentQuery.FieldByName("TABTIPOLEIAUTE").AsInteger = 1 Then
  bsShowMessage("Leiaute definido nos parâmetros gerais é variável!", "E")
  CanContinue = False
  Exit Sub
ElseIf Q.FieldByName("TABTIPOLEIAUTE").AsInteger <> 4 And CurrentQuery.FieldByName("TABTIPOLEIAUTE").AsInteger = 2 Then
  bsShowMessage("Leiaute definido nos parâmetros gerais é fixo!","E")
  CanContinue = False
  Exit Sub
End If

Set Q = Nothing
End Sub

Public Sub TABLE_NewRecord()
'SMS 52120 - Marcelo Barbosa - 26/12/2005
  If (VisibleMode And NodeInternalCode = 500) Or (WebMode And WebMenuCode = "T5140") Then
    If VisibleMode Then
      TABTIPOROTINA.Pages(1).Visible = True
      TABTIPOROTINA.Pages(0).Visible = False
    End If
    CurrentQuery.FieldByName("TABTIPOROTINA").Value = 2

    BOTAOAGENDAR.Visible = False
    BOTAOCANCELAFATURAMENTO.Visible = False
    BOTAOCANCELAR.Visible = False
    BOTAOCOMUNICADO.Visible = False
    BOTAODESBLOQUEAR.Visible = False
    BOTAOEXPORTAR.Visible = False
    BOTAOFATURA.Visible = False
    BOTAOGERAR.Visible = False
    BOTAOIMPRIMIR.Visible = False
    BOTAOPROTOCOLO.Visible = False

  Else
    If VisibleMode Then
      TABTIPOROTINA.Pages(0).Visible = True
      TABTIPOROTINA.Pages(1).Visible = False
    End If
    CurrentQuery.FieldByName("TABTIPOROTINA").Value = 1
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	 Select Case CommandID
      Case "BOTAOAGENDAR"
		BOTAOAGENDAR_OnClick
      Case "BOTAOCANCELAFATURAMENTO"
        BOTAOCANCELAFATURAMENTO_OnClick
      Case "BOTAOCANCELAR"
        BOTAOCANCELAR_OnClick
      Case "BOTAOCOMUNICADO"
        BOTAOCOMUNICADO_OnClick
      Case "BOTAODESBLOQUEAR"
		BOTAODESBLOQUEAR_OnClick
      Case "BOTAOEXPORTAR"
        BOTAOEXPORTAR_OnClick
      Case "BOTAOFATURA"
        BOTAOFATURA_OnClick
      Case "BOTAOGERAR"
        BOTAOGERAR_OnClick
      Case "BOTAOIMPRIMIR"
		BOTAOIMPRIMIR_OnClick
      Case "BOTAOPROTOCOLO"
        BOTAOPROTOCOLO_OnClick
      Case "BOTAORETORNO"
        BOTAORETORNO_OnClick
  End Select
End Sub


