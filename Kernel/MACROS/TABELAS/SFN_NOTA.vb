'HASH: 1CA87C93B6D4813199D18E326FCAE767

'Macro SFN_NOTA -Keila
'#Uses "*bsShowMessage"
'#Uses "*ProcuraPrestador"

Option Explicit
Dim vHandle As Long
Dim Duplicado As Boolean


Public Sub BOTAOCANCCONCILIACAO_OnClick()

  Dim UpDoc As Object
  Dim QDoc As Object
  Dim qDocBx As Object
  Set QDoc = NewQuery
  Set UpDoc = NewQuery
  Set qDocBx = NewQuery

  Dim QNota As Object
  Set QNota = NewQuery

  If bsshowmessage("Deseja realmente cancelar essa conciliação?", "Q") = vbNo Then
     Exit Sub
  End If

  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    QNota.Add("SELECT HANDLE FROM SFN_NOTA")
    QNota.Add("WHERE HANDLE IN (SELECT NOTA FROM SFN_NOTA_DOCUMENTO WHERE NOTA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
    QNota.Active = True

    If QNota.EOF Then
      BSShowMessage("Esta Nota não está conciliada", "I")
      Exit Sub
    End If
  End If

  qDocBx.Clear
  qDocBx.Add("SELECT COUNT(1) QTDE   ")
  qDocBx.Add("  FROM SFN_DOCUMENTO D ")
  qDocBx.Add("  JOIN SFN_NOTA_DOCUMENTO ND ON D.HANDLE = ND.DOCUMENTO")
  qDocBx.Add(" WHERE ND.NOTA =:HNOTA ")
  qDocBx.Add("   AND D.BAIXADATA IS NOT NULL ")
  qDocBx.ParamByName("HNOTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qDocBx.Active = True

  If qDocBx.FieldByName("QTDE").AsInteger > 0 Then
    If bsShowMessage("Existe documento baixado vinculado a esta nota. Deseja continuar?", "Q") = vbNo Then
      Exit Sub
    End If
  End If

  QDoc.Add("SELECT DOCUMENTO FROM SFN_NOTA_DOCUMENTO")
  QDoc.Add("WHERE NOTA =:HNOTA ")
  QDoc.ParamByName("HNOTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QDoc.Active = True

  If Not QDoc.EOF Then
    'Pinheiro - sms 61822
    UpDoc.Clear
	If Not InTransaction Then StartTransaction
	    UpDoc.Add("UPDATE SFN_DOCUMENTO SET LIBERACAODATA = NULL              ")
	    UpDoc.Add(" WHERE HANDLE IN (SELECT DOCUMENTO FROM SFN_NOTA_DOCUMENTO ")
	    UpDoc.Add("                   WHERE NOTA =:HNOTA)                     ")
	    UpDoc.ParamByName("HNOTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	    UpDoc.ExecSQL

    	UpDoc.Clear
    	UpDoc.Add("DELETE FROM SFN_NOTA_DOCUMENTO")  ' SMS 95929 - Paulo Melo - 18/04/2008 - Delete estava sem FROM, isso dá problema em DB2
    	UpDoc.Add(" WHERE NOTA =:HNOTA      ")
    	UpDoc.ParamByName("HNOTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    	UpDoc.ExecSQL
    If InTransaction Then Commit
    BSShowMessage("Conciliamento Cancelado", "I")
  Else
    BSShowMessage("Esta nota não foi conciliada", "I")

  End If
  RefreshNodesWithTable("SFN_NOTA")


End Sub

Public Sub BOTAOCANCELA_OnClick()
  Dim vErro As String
  Dim QTipoNota As Object
  Dim QCanc As Object
  Dim QCancNota As Object
  Dim Interface As Object
  Dim vLog As String

  Set QTipoNota = NewQuery
  Set QCanc = NewQuery
  Set QCancNota = NewQuery
  Set Interface = NewQuery


  vErro = ""
  'vErro =VerificaDoc
  'If vErro <>"" Then
  ' MsgBox "Não foi possível cancelar NOta fiscal: " +vErro
  'Else
  'If CurrentQuery.FieldByName("NOTAIMPRESSA").AsString ="N" Then'
  '   QTipoNota.Add("SELECT TN.CANCELADOCUMENTO ")
  '  QTipoNota.Add("  FROM SFN_TIPONOTA TN,              ")
  ' QTipoNota.Add("       SFN_NOTA N                    ")
  '      QTipoNota.Add(" WHERE TN.HANDLE = N.TIPO            ")
  '     QTipoNota.Add("   AND N.HANDLE = :PNOTA             ")
  '    QTipoNota.ParamByName("PNOTA").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  '   QTipoNota.Active =True

  '      QCanc.Add("SELECT DOCUMENTO FROM SFN_NOTA_DOCUMENTO    ")
  '     QCanc.Add("  WHERE NOTA = :PNOTA                       ")
  '    QCanc.ParamByName("PNOTA").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  '   QCanc.Active =True
  '
  '      If QTipoNota.FieldByName("CANCELADOCUMENTO").AsString ="S" Then
  Set Interface = CreateBennerObject("SFNCANCEL.CANCELAMENTO")
  Interface.CancelaNota(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing
  '      End If

  '      DesvinculaNota
  '
  '     QCancNota.Add("UPDATE SFN_NOTA                   ")
  '    QCancNota.Add("  SET DATACANCELAMENTO = :PDATA   ")
  '   QCancNota.Add("WHERE HANDLE = :PNOTA             ")
  '  QCancNota.ParamByName("PNOTA").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  ' QCancNota.ParamByName("PDATA").Value =ServerDate
  'QCancNota.ExecSQL

  '      vLog ="Cancela Nota fiscal" +Chr(13)
  '     vLog =vLog +"Data Cancelamento: " +CurrentQuery.FieldByName("DATACANCELAMENTO").AsString +Chr(13)
  '    WriteAudit("C",HandleOfTable("SFN_NOTA"),CurrentQuery.FieldByName("handle").AsInteger,vLog)

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  '    Else
  '     MsgBox "Nota já foi impressa não pode ser cancelada!"
  '   End If
  'End If
  Set QTipoNota = Nothing
  Set QCanc = Nothing
  Set QCancNota = Nothing

End Sub

Public Sub BOTAOCONCILIAR_OnClick()
  Dim Interface As Object
  Set Interface = CreateBennerObject("SFNNota.Rotinas")

  If CurrentQuery.State <>1 Then
    MsgBox "Salve ou Cancele as alterações antes de conciliar a nota."
    Exit Sub
  End If

  Dim QNota As Object
  Set QNota = NewQuery

  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    QNota.Add("SELECT HANDLE FROM SFN_NOTA")
    QNota.Add("WHERE HANDLE IN (SELECT NOTA FROM SFN_NOTA_DOCUMENTO WHERE NOTA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
    QNota.Active = True

    If Not QNota.EOF Then
      If VisibleMode Then
        BSShowMessage("Esta Nota está conciliada", "I")
        Set QNota = Nothing
        Exit Sub
      End If
    End If

  End If

  If Not CurrentQuery.State = 3 Then
     Set Interface = CreateBennerObject("BSINTERFACE0043.RotinaNota")
     Interface.Conciliar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "C")

     CurrentQuery.Active = False
     CurrentQuery.Active = True
     RefreshNodesWithTable("SFN_NOTA")

     Set Interface = Nothing
  Else
    BSShowMessage("Confirme a inclusão antes de conciliar a nota.", "I")
  End If
  Set QNota = Nothing

End Sub

Public Sub BOTAOROTINACONCILIACAO_OnClick()
  Dim Interface As Object

  Set Interface = CreateBennerObject("SFNNota.Rotinas")
  Interface.Conciliar(CurrentSystem, 0, "N")
  Set Interface = Nothing

End Sub

Public Sub BOTAOEXCLUIR_OnClick()
  Dim vAux As Long
  Dim vErro As String
  Dim vMsg As String
  Dim QExcluiNOta As Object
  Dim QAvulsa As Object
  Dim qvinculo As Object
  Dim Interface As Object
  Dim Inter As Object
  Dim ExcluiDLL As Object
  Dim ExcluiItem As Object

  Set QExcluiNOta = NewQuery
  Set QAvulsa = NewQuery
  Set qvinculo = NewQuery
  Set ExcluiItem = NewQuery

  If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    vErro = ""
    vErro = VerificaDoc
    If vErro <>"" Then
      vMsg = "Não foi possível excluir nota fiscal: " + vErro
    Else
      If CurrentQuery.FieldByName("NOTAIMPRESSA").AsString = "N" Then
        QAvulsa.Add("SELECT NOTAAVULSA FROM SFN_TIPONOTA WHERE HANDLE = " + CurrentQuery.FieldByName("TIPO").AsString)
        QAvulsa.Active = True

        If QAvulsa.FieldByName("NOTAAVULSA").AsString = "S" Then
          qvinculo.Add("SELECT NF.FATURA, ND.DOCUMENTO")
          qvinculo.Add("FROM SFN_NOTA_FATURA NF,")
          qvinculo.Add("     SFN_NOTA_DOCUMENTO ND,")
          qvinculo.Add("     SFN_NOTA N")
          qvinculo.Add("WHERE N.HANDLE = ND.NOTA")
          qvinculo.Add("  AND N.HANDLE = NF.NOTA")
          qvinculo.Add("  AND N.HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString)
          qvinculo.Active = True

          Set Interface = CreateBennerObject("FINANCEIRO.DOCUMENTO")
          Set Inter = CreateBennerObject("FINANCEIRO.FATURA")
          vAux = Interface.Excluir(CurrentSystem, qvinculo.FieldByName("DOCUMENTO").AsInteger)
          If vAux >0 Then
            While Not qvinculo.EOF
              vAux = Inter.ExcluirFatura(CurrentSystem, qvinculo.FieldByName("FATURA").AsInteger)
              qvinculo.Next
            Wend
          End If
          vMsg = "Exclusão de Nota fiscal, fatura e Documento bem sucedida"
        Else
          DesvinculaNota
        End If
        Set ExcluiDLL = CreateBennerObject("SFNNOTA.Rotinas")
        ExcluiDLL.Excluir(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vAux)
        If vAux >0 Then
          If vMsg = "" Then
            vMsg = "Nota Fiscal excluída com Sucesso!"
          End If
        Else
          vMsg = "Erro na exclusão da Nota fiscal"
        End If
      Else
        vMsg = "Nota já foi impressa não pode ser Excluída!"
      End If
    End If
  Else
    vMsg = "Nota fiscal está cancelada!"
  End If
  If vMsg <>"" Then
    bsShowMessage(vMsg, "I")
  End If
  Set QExcluiNOta = Nothing
  Set QAvulsa = Nothing
  Set qvinculo = Nothing
  Set Interface = Nothing
  Set Inter = Nothing
  Set ExcluiDLL = Nothing
  Set ExcluiItem = Nothing
  RefreshNodesWithTable("SFN_NOTA")
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("NOTAIMPRESSA").AsString = "S" Then
    MsgBox("Esta nota já foi impressa. Executar a rotina de reimpressão")
    Exit Sub
  End If

  Dim Obj As Object

  Set Obj = CreateBennerObject("SamImpressao.NotaFiscal")
  Obj.Inicializar(CurrentSystem)
  Obj.ImprimirNota(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Obj.Finalizar
  Set Obj = Nothing

  WriteAudit("I", HandleOfTable("SFN_NOTA"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Impressão")

End Sub

Public Sub BOTAOPROCURAPESSOA_OnClick()
'  Dim Interface As Object
'  Dim vHandle As Long
'  Dim SQL As Object

'  Set SQL = NewQuery
'  Set Interface = CreateBennerObject("Procura.Procurar")
'  vHandle = Interface.Exec(CurrentSystem, "SFN_PESSOA", "CNPJCPF|NOME", 2, "CNPJ/CPF|NOme", "HANDLE > 0", "Procura Pessoa", True, "")

End Sub

Public Sub BOTAOPROCURAPRESTADOR_OnClick()
 ' Dim vHandle As Long
 ' Dim SQL As Object


'  Set SQL = NewQuery


'  vHandle = ProcuraPrestador("C", "T", "")' pelo CPF e todos
'  SQL.Add("SELECT handle FROM SFN_CONTAFIN WHERE PRESTADOR = :PHANDLE")
'  SQL.ParamByName("PHANDLE").Value = vHandle
'  SQL.Active = True
'  If SQL.FieldByName("Handle").AsInteger <>0 Then
'    CurrentQuery.Edit
'    CurrentQuery.FieldByName("CONTAFINANCEIRA").Value = SQL.FieldByName("Handle").Value
'  End If
'  If Not CurrentQuery.FieldByName("CONTAFINANCEIRA").IsNull Then
'    AtualizaRotulo
'  End If
End Sub

Public Sub PESSOA_OnExit()
    'Dim SQL As Object

    'Set SQL = NewQuery
    'SQL.Clear
    'SQL.Add("SELECT handle FROM SFN_CONTAFIN WHERE PESSOA = :PHANDLE")
	'SQL.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("PESSOA").AsInteger
	'SQL.Active = True
	'If SQL.FieldByName("Handle").AsInteger <>0 Then
	'  CurrentQuery.Edit
	'  CurrentQuery.FieldByName("CONTAFINANCEIRA").Value = SQL.FieldByName("Handle").Value
	'End If

	'If Not CurrentQuery.FieldByName("CONTAFINANCEIRA").IsNull Then
	'  AtualizaRotulo
	'End If
	'Set SQL = Nothing

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  '#Uses "*ProcuraPrestador"
  Dim vHandle As Long
  'Dim SQL As Object
  Dim vDigitado As String

  'Set SQL = NewQuery

  If Len(PRESTADOR.LocateText) > 0 Then
    vDigitado = PRESTADOR.LocateText
  Else
    vDigitado = ""
  End If

  vHandle = ProcuraPrestador("C", "T", vDigitado)' pelo CPF e todos

  If vHandle <> 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If

  'Set SQL = NewQuery
  'SQL.Add("SELECT handle FROM SFN_CONTAFIN WHERE PRESTADOR = :PHANDLE")
  'SQL.ParamByName("PHANDLE").Value = vHandle
  'SQL.Active = True
  'If SQL.FieldByName("Handle").AsInteger <>0 Then
  '  CurrentQuery.Edit
  '  CurrentQuery.FieldByName("CONTAFINANCEIRA").Value = SQL.FieldByName("Handle").Value
  'End If
  'If Not CurrentQuery.FieldByName("CONTAFINANCEIRA").IsNull Then
  '  AtualizaRotulo
  'End If

  'Set SQL = Nothing
End Sub

Public Sub TABLE_AfterCommitted()

   If VisibleMode Then
	   Dim Interface As Object

	   Set Interface = CreateBennerObject("BSINTERFACE0043.RotinaNota")
	   Interface.Conciliar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "C")

	   CurrentQuery.Active = False
	   CurrentQuery.Active = True

	   Set Interface = Nothing
	End If

End Sub


Public Sub TABLE_AfterScroll()
  Dim Rotulo As String
  Dim QNota As Object
  Set QNota = NewQuery
  Dim SQL As Object
  Set SQL = NewQuery
  Dim vQtde As Integer

  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    SessionVar("HNOTA") = CurrentQuery.FieldByName("HANDLE").AsString

    QNota.Add("SELECT HANDLE FROM SFN_NOTA")
    QNota.Add("WHERE HANDLE IN (SELECT NOTA FROM SFN_NOTA_DOCUMENTO WHERE NOTA = " + CurrentQuery.FieldByName("HANDLE").AsString + ")")
    QNota.Active = True

    vQtde = 0
    SQL.Clear
    SQL.Add("SELECT COUNT(1) QTDE      ")
    SQL.Add("  FROM SFN_NOTA_DOCUMENTO ")
    SQL.Add(" WHERE NOTA = " + CurrentQuery.FieldByName("HANDLE").AsString)
    SQL.Active = True
    vQtde = SQL.FieldByName("QTDE").AsInteger

    SQL.Clear
    SQL.Add("SELECT D.NUMERO, ")
    SQL.Add("       D.DATAVENCIMENTO, ")
    SQL.Add("       D.VALOR ")
    SQL.Add("  FROM SFN_DOCUMENTO D, ")
    SQL.Add("       SFN_NOTA_DOCUMENTO ND")
    SQL.Add(" WHERE ND.DOCUMENTO = D.HANDLE")
    SQL.Add("   AND ND.NOTA = " + CurrentQuery.FieldByName("HANDLE").AsString)
    SQL.Active = True

    If Not QNota.EOF Then
      If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 2 Then
        Rotulo = "Documento nº: " + SQL.FieldByName("NUMERO").AsString + "     Vencimento: " + SQL.FieldByName("DATAVENCIMENTO").AsString
        ROTULOBENEFICIARIO.Text = Rotulo
      End If
      If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 1 Then
        If vQtde > 1 Then
          ROTULODOCUMENTO.Text = ""
        Else
          Rotulo = "Documento nº: " + SQL.FieldByName("NUMERO").AsString + "     Vencimento: " + SQL.FieldByName("DATAVENCIMENTO").AsString
          ROTULODOCUMENTO.Text = Rotulo
        End If
      End If
    Else
      ROTULOBENEFICIARIO.Text = ""
      ROTULODOCUMENTO.Text = ""
    End If
    AtualizaRotulo
  End If
  Set QNota = Nothing
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  ROTULOPRESTADOR.Text = ""
  ROTULOBENEFICIARIO.Text = ""
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim TIPO As Object
  Set TIPO = NewQuery
  Dim SQL As Object

  TIPO.Add("SELECT TABORIGEM, OBRIGATORIEDADESERIE FROM SFN_TIPONOTA WHERE HANDLE = :HANDLETIPO")
  TIPO.ParamByName("HANDLETIPO").AsString = CurrentQuery.FieldByName("TIPO").AsString
  TIPO.Active = True

  If TIPO.FieldByName("TABORIGEM").AsInteger <>CurrentQuery.FieldByName("TABORIGEM").AsInteger Then
    bsShowMessage("Origem incompatível com o tipo de Nota", "E")
    CanContinue = False
  End If

  If TIPO.FieldByName("OBRIGATORIEDADESERIE").AsString = "S" Then
  	If CurrentQuery.FieldByName("SERIE").IsNull Then
	  BsShowMessage("Obrigatório o preenchimento do campo ''Série'' para esse tipo de nota!", "E")
      CanContinue = False
  	End If
  End If

  If CurrentQuery.FieldByName("TABORIGEM").AsInteger = 1 Then
    Set SQL = NewQuery
	If CurrentQuery.FieldByName("TABPRESTADORTIPO").AsInteger = 1 Then
	  SQL.Add("SELECT handle FROM SFN_CONTAFIN WHERE PRESTADOR = :PHANDLE")
	  SQL.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	  SQL.Active = True
	  If SQL.FieldByName("Handle").AsInteger <>0 Then
	    CurrentQuery.Edit
	    CurrentQuery.FieldByName("CONTAFINANCEIRA").Value = SQL.FieldByName("Handle").Value
	  End If
	  If Not CurrentQuery.FieldByName("CONTAFINANCEIRA").IsNull Then
	    AtualizaRotulo
	  End If
	Else
	  SQL.Clear
	  SQL.Add("SELECT handle FROM SFN_CONTAFIN WHERE PESSOA = :PHANDLE")
	  SQL.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("PESSOA").AsInteger
	  SQL.Active = True
	  If SQL.FieldByName("Handle").AsInteger <>0 Then
	    CurrentQuery.Edit
	    CurrentQuery.FieldByName("CONTAFINANCEIRA").Value = SQL.FieldByName("Handle").Value
	  End If
	  If Not CurrentQuery.FieldByName("CONTAFINANCEIRA").IsNull Then
	    AtualizaRotulo
	  End If
	End If
    Set SQL = Nothing
  End If

  Dim q As Object
  Set q = NewQuery
  q.Active = False
  q.Clear
  q.Add("  SELECT HANDLE, COUNT(*) N")
  q.Add("    FROM SFN_NOTA")
  If CurrentQuery.FieldByName("NUMERO").IsNull Then
    q.Add("   WHERE NUMERO = NULL")
  Else
    q.Add("   WHERE NUMERO = :NUMERO")
    q.ParamByName("NUMERO").AsString = CurrentQuery.FieldByName("NUMERO").AsString
  End If
  q.Add("     AND CONTAFINANCEIRA = :CONTAFIN")
  q.Add("     AND SERIE = :SERIE             ")
  q.Add("GROUP BY HANDLE                     ")
  q.ParamByName("CONTAFIN").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
  q.ParamByName("SERIE").AsInteger = CurrentQuery.FieldByName("SERIE").AsInteger
  q.Active = True
  If (q.FieldByName("N").AsInteger <> 0) And (q.FieldByName("HANDLE").AsInteger <> CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("Já existe uma Nota com o mesmo Número e Série para este Prestador!", "E")
    CanContinue = False
  End If
  Set q = Nothing
End Sub

Public Sub TABLE_NewRecord()
  If VisibleMode Then
    If NodeInternalCode <> 721 Then
      TABORIGEM.Pages(1).Visible = False
    End If
    If Len(ROTULODOCUMENTO.Text) > 0 Then
      ROTULODOCUMENTO.Text  = ""
    End If
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
    Select Case CommandID
      Case "BOTAOCANCCONCILIACAO"
        BOTAOCANCCONCILIACAO_OnClick
    End Select
End Sub

Public Sub TIPO_OnChange()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT * FROM SFN_TIPONOTA")
  SQL.Add("WHERE HANDLE = :PTIPO")
  SQL.ParamByName("PTIPO").AsInteger = CurrentQuery.FieldByName("TIPO").AsInteger
  SQL.Active = True

  Set SQL = Nothing
End Sub

Function VerificaDoc()As String
  Dim QBAIXA As Object
  Dim QRotArq As Object
  Set QRotArq = NewQuery
  Set QBAIXA = NewQuery

  QBAIXA.Add("SELECT D.HANDLE, D.BAIXADATA          ")
  QBAIXA.Add("  FROM SFN_NOTA_DOCUMENTO ND,         ")
  QBAIXA.Add("       SFN_DOCUMENTO D                ")
  QBAIXA.Add("WHERE ND.NOTA = :PNota                ")
  QBAIXA.Add("  AND D.HANDLE = ND.DOCUMENTO         ")
  QBAIXA.ParamByName("PNOTA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  QBAIXA.Active = True

  If QBAIXA.FieldByName("BAIXADATA").Value = Null Then
    VerificaDoc = "Documento Baixado"
  End If

  QRotArq.Add("SELECT RA.ROTINAARQUIVO          ")
  QRotArq.Add("  FROM SFN_DOCUMENTO D,          ")
  QRotArq.Add("       SFN_ROTINAARQUIVO_DOC RA  ")
  QRotArq.Add("WHERE D.HANDLE = RA.DOCUMENTO    ")
  QRotArq.Add("  AND D.HANDLE = :PDOC           ")
  QRotArq.ParamByName("PDOC").Value = QBAIXA.FieldByName("HANDLE").AsInteger
  QRotArq.Active = True

  If Not QRotArq.EOF Then
    VerificaDoc = "Documento em Rotina Arquivo"
  End If
  Set QBAIXA = Nothing
  Set QRotArq = Nothing
End Function

Public Sub DesvinculaNota()
  Dim QDoc As Object
  Dim QFat As Object
  Set QDoc = NewQuery
  Set QFat = NewQuery

  QDoc.Add("DELETE SFN_NOTA_DOCUMENTO")
  QDoc.Add("WHERE NOTA = " + CurrentQuery.FieldByName("HANDLE").AsString)
  QDoc.ExecSQL

  QFat.Add("DELETE SFN_NOTA_FATURA")
  QFat.Add("WHERE NOTA =" + CurrentQuery.FieldByName("HANDLE").AsString)
  QFat.ExecSQL

  Set QDoc = Nothing
  Set QFat = Nothing
End Sub


Public Sub AtualizaRotulo
  Dim Rotulo As String

  Dim QRotulo As Object
  Set QRotulo = NewQuery

  Dim QDado As Object
  Set QDado = NewQuery


  If Not CurrentQuery.FieldByName("HANDLE").IsNull Or CurrentQuery.State = 3 Then
    QRotulo.Add("SELECT PRESTADOR, PESSOA FROM SFN_CONTAFIN WHERE HANDLE = :CONTAFINANCEIRA")
    QRotulo.ParamByName("CONTAFINANCEIRA").AsString = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsString
    QRotulo.Active = True

    If Not QRotulo.FieldByName("PRESTADOR").IsNull Then
      QDado.Clear
      QDado.Add("SELECT NOME, CPFCNPJ FROM SAM_PRESTADOR WHERE HANDLE = :PREST")
      QDado.ParamByName("PREST").Value = QRotulo.FieldByName("PRESTADOR").AsInteger
      QDado.Active = True

      Rotulo = "Prestador: " + QDado.FieldByName("NOME").AsString
      ROTULOPRESTADOR.Text = Rotulo + "    CNPJ/CPF: " + QDado.FieldByName("CPFCNPJ").AsString
    Else
      QDado.Clear
      QDado.Add("SELECT NOME, CNPJCPF FROM SFN_PESSOA WHERE HANDLE = :PESSOA")
      QDado.ParamByName("PESSOA").Value = QRotulo.FieldByName("PESSOA").AsInteger
      QDado.Active = True

      Rotulo = "Pessoa: " + QDado.FieldByName("NOME").AsString
      ROTULOPRESTADOR.Text = Rotulo + "    CNPJ/CPF: " + QDado.FieldByName("CNPJCPF").AsString
    End If
  End If
  Set QDado = Nothing
  Set QRotulo = Nothing
End Sub
