'HASH: 8E4B3F84BE1092124C2375956319F385
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOFECHARCOTACAO_OnClick()
  Dim sp As Object
  Set sp = NewStoredProc

  sp.Name = "BSLIC_FECHAMENTO"
  sp.AutoMode = True
  sp.AddParam("P_HANDLELICITACAO",ptInput)
  sp.ParamByName("P_HANDLELICITACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sp.ExecProc

  Set sp = Nothing

  Dim vHandleProcessoAgendado As Long

  vHandleProcessoAgendado = CurrentQuery.FieldByName("PROCESSOAGENDADO").AsInteger

  If Not InTransaction Then
    StartTransaction
  End If

  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("UPDATE SAM_LICITACAO SET PROCESSOAGENDADO = NULL")
  sql.Add("WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL

  Set sql = Nothing

  If vHandleProcessoAgendado > 0 Then
    Dim obj As Object
    Set obj = CreateBennerObject("BSPORTALWEB.WEB")
    obj.ApagarAgendamentoLicitacao(CurrentSystem,vHandleProcessoAgendado)
    Set obj = Nothing
  End If

  RefreshNodesWithTable("SAM_LICITACAO")

  If InTransaction Then
    Commit
  End If

  bsShowMessage("Fechamento concluído.", "I")

End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("ANO").AsDateTime = ServerDate
  Dim vDataHora As Date
  vDataHora = ServerNow

  Dim vSec As Integer
  vSec = DatePart("s",vDataHora)

  CurrentQuery.FieldByName("DATAHORAABERTURA").AsDateTime = DateAdd("s",-vSec, DateAdd("n",10,vDataHora))


  Dim vNumero As Long
  NewCounter("SAM_LICITACAO_NUMERO",Year(ServerDate),1,vNumero)
  CurrentQuery.FieldByName("NUMERO").AsInteger = vNumero

End Sub
Public Function VerificaLanceMaiorPrecoMaximo(pFornecedor As Long, pLicitacao As Long) As String
  Dim sql As Object
  Set sql = NewQuery

  VerificaLanceMaiorPrecoMaximo = ""

  sql.Clear
  sql.Add("  SELECT E.DESCRICAO PRODUTO, ")
  sql.Add("         A. PRECOUNITARIO, ")
  sql.Add("         D. PRECOMAXIMO  ,")
  sql.Add("         F.DESCRICAO FORNECEDOR,")
  sql.Add("         G.DESCRICAO APRESENTACAO")
  sql.Add("    FROM SAM_LICITACAO_FORNECEDOR_ITEM A")
  sql.Add("    JOIN SAM_LICITACAO_FORNECEDOR      B ON (A.LICITACAOFORNECEDOR = B.HANDLE)")
  sql.Add("    JOIN SAM_LICITACAO                 C ON (B.LICITACAO = C.HANDLE)")
  sql.Add("    JOIN SAM_LICITACAO_ITENS           D ON (D.LICITACAO = C.HANDLE and A.ITEM = D.HANDLE)")
  SQL.Add("    JOIN SAM_MATMED                    E ON (D.MATMED = E.HANDLE)")
  SQL.Add("    LEFT JOIN SAM_MATMEDBRFORNECEDOR   F ON (E.BRFORNECEDOR = F.HANDLE)")
  SQL.Add("    LEFT JOIN SAM_MATMEDBRAPRESENTACAO G ON (E.BRAPRESENTACAO = G.HANDLE)")
  SQL.Add("   WHERE B.FORNECEDOR = :FORNECEDOR")
  SQL.Add("     AND C.HANDLE     = :LICITACAO")
  SQL.Add("     AND D.PRECOMAXIMO < A.PRECOUNITARIO")
  SQL.ParamByName("FORNECEDOR").AsInteger = pFornecedor
  SQL.ParamByName("LICITACAO").AsInteger = pLicitacao
  SQL.Active = True
  If Not SQL.EOF Then
    VerificaLanceMaiorPrecoMaximo = "Valores cotados acima do preço máximo estabelecido. Os seguintes produtos devem ter os preços alterados: " +Chr(13)+Chr(10)
  End If

  While Not SQL.EOF
    VerificaLanceMaiorPrecoMaximo = VerificaLanceMaiorPrecoMaximo + "Produto: "+SQL.FieldByName("PRODUTO").AsString + Chr(13)+Chr(10)+ _
    "Fornecedor  : " + SQL.FieldByName("FORNECEDOR").AsString + Chr(13) + Chr(10) + _
    "Apresentação: " + SQL.FieldByName("APRESENTACAO").AsString + Chr(13) + Chr(10) + _
    "Preço máximo: " + Format(SQL.FieldByName("PRECOMAXIMO").AsString, "###,###,##0.00") + Chr(13) + Chr(10) + Chr(13) + Chr(10)
    SQL.Next
  Wend



End Function



Public Sub EnviaEmailFornecedor(pLicitacao As Integer)
  Dim sql As Object
  Dim sql2 As Object
  Dim vNomeRelatorio As String


  Set sql = NewQuery
  Set sql2 = NewQuery

  sql.Clear
  sql.Add("SELECT B.EMAILRESPONSAVEL, C.ANO, C.NUMERO ")
  sql.Add("  FROM SAM_LICITACAO_FORNECEDOR A")
  sql.Add("  JOIN SFN_PESSOA               B ON (A.FORNECEDOR = B.HANDLE)")
  sql.Add("  JOIN SAM_LICITACAO            C ON (A.LICITACAO = C.HANDLE) ")
  sql.Add("  WHERE LICITACAO = :LICITACAO")
  sql.ParamByName("LICITACAO").AsInteger = pLicitacao
  sql.Active = True

  sql2.Clear
  sql2.Add("SELECT RELATORIOCOTACAO FROM SAM_PARAMETROSWEB")
  sql2.Active = True

  vNomeRelatorio = "Cotacao_" + Trim(Str(Year(sql.FieldByName("ANO").AsDateTime))+"_"+sql.FieldByName("NUMERO").AsString) + ".pdf"

  While Not sql.EOF
    If Not sql.FieldByName("EMAILRESPONSAVEL").IsNull Then
      ReportExport(sql2.FieldByName("RELATORIOCOTACAO").AsInteger,"A.LICITACAO = " + CurrentQuery.FieldByName("HANDLE").AsString, vNomeRelatorio,False,False,sql.FieldByName("EMAILRESPONSAVEL").AsString,"")
    End If
    sql.Next
  Wend

  Set sql2 = Nothing
  Set sql = Nothing
End Sub



Public Function LocalizaFornecedor(pUsuario As Long) As Long
  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT PESSOA FROM Z_GRUPOUSUARIOS_PESSOA WHERE USUARIO = :USUARIO")
  sql.ParamByName("USUARIO").AsInteger = pUsuario
  sql.Active = True
  LocalizaFornecedor = sql.FieldByName("PESSOA").AsInteger
  Set sql = Nothing
End Function


Public Function AnexarArquivo(pComplementoNomeArquivo As String) As String
  Dim o As Object
  Dim HandleMsg, HandleAnexo As Long
  Dim NomeArq, NomeArqServer As String
  Dim QSit As Object

  If CurrentQuery.State <> 1 Then
    'MsgBox("Operação Cancelada. Registro está em Edição",vbCritical)
    BsShowMessage("Operação Cancelada. Registro está em Edição", "I")
    Exit Function
  End If

 NomeArq = OpenDialog
 NomeArqServer = NomeArq
  While InStr(NomeArqServer,"\") <> 0
    NomeArqServer = Mid(NomeArqServer, InStr(NomeArqServer,"\") + 1, Len(NomeArqServer))
  Wend

 If NomeArq <> "" Then
  Dim OBJ As Object
  Set OBJ = SuperServerClient("DOC")
  OBJ.Select("EDITAL")

'   HandleAnexo = CurrentQuery.FieldByName("HANDLE").AsInteger

  OBJ.SetDocument(NomeArq,"ARQUIVO_"+ pComplementoNomeArquivo+".BDF")
  OBJ.Select("")
  Set OBJ = Nothing
 End If

 AnexarArquivo = NomeArqServer
End Function


Public Sub BOTAOABRIRAUTORIZACAO_OnClick()
  Dim Interface As Object
  Dim vAutorizacao As Long

  UserVar("BENEFICIARIO") =""
  Set Interface =CreateBennerObject("CA043.Autorizacao")
  vAutorizacao = Interface.Executar(CurrentSystem,0,0,0)
  Set Interface =Nothing
  If vAutorizacao > 0 Then
    If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
      CurrentQuery.FieldByName("AUTORIZACAO").AsInteger = vAutorizacao
    Else
      CurrentQuery.Edit
      CurrentQuery.FieldByName("AUTORIZACAO").AsInteger = vAutorizacao
      CurrentQuery.Post
    End If
  End If
End Sub

Public Sub BOTAOABRIRLICITACAO_OnClick()
  Dim DLL As Object
  Dim sql As Object
  Set sql = NewQuery
  Dim vDataAgendamento As Date

  If Not InTransaction Then
    StartTransaction
  End If

  If (CurrentQuery.FieldByName("PUBLICADAWEB").AsString = "N") And (CurrentQuery.State = 1) Then
    Set DLL = CreateBennerObject("BSPORTALWEB.WEB")
    DLL.SelecionarFornecedor(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,0)
    vDataAgendamento = DateAdd("n",5,CurrentQuery.FieldByName("DATAENCERRAMENTO").AsDateTime) '5 minutos a mais pra rodar o fechamento
    DLL.AgendarFecharLicitacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,vDataAgendamento)
    sql.Clear
    sql.Add("UPDATE SAM_LICITACAO SET PUBLICADAWEB = 'S' WHERE HANDLE = :HANDLE")
    sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ExecSQL
    EnviaEmailFornecedor(CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set sql = Nothing
    Set DLL = Nothing
  Else
    If CurrentQuery.State <> 1 Then
      BsShowMessage("Registro está em edição. Operação não permitida.", "I")
    Else
      BsShowMessage("Cotação já está iniciada.", "I")
    End If

  End If

  If InTransaction Then
    Commit
  End If
End Sub

Public Sub BOTAOANEXAREDITALCOMPLETO_OnClick()
  Dim vAnexo As String
  Dim sql As Object

  vAnexo = AnexarArquivo(CurrentQuery.FieldByName("HANDLE").AsString)

  If InTransaction Then
    StartTransaction
  End If

  Set sql = NewQuery
  sql.Clear
  sql.Add("UPDATE SAM_LICITACAO SET EDITALCOMPLETO = :EDITAL WHERE HANDLE = :HANDLE")
  sql.ParamByName("EDITAL").AsString = vAnexo
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL

  RefreshNodesWithTable("SAM_LICITACAO")

  If InTransaction Then
    Commit
  End If


  Set sql = Nothing

End Sub

Public Sub BOTAOANEXAREDITALRESUMIDO_OnClick()
  Dim vAnexo As String
  Dim sql As Object

  vAnexo = AnexarArquivo(CurrentQuery.FieldByName("HANDLE").AsString+"_1")

  If InTransaction Then
    StartTransaction
  End If


  Set sql = NewQuery
  sql.Clear
  sql.Add("UPDATE SAM_LICITACAO SET EDITALRESUMIDO = :EDITAL WHERE HANDLE = :HANDLE")
  sql.ParamByName("EDITAL").AsString = vAnexo
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL

  If InTransaction Then
    Commit
  End If


  RefreshNodesWithTable("SAM_LICITACAO")

  Set sql = Nothing

End Sub

Public Sub BOTAODETAUTORIZACAO_OnClick()
  If CurrentQuery.FieldByName("AUTORIZACAO").AsInteger > 0 Then
    Dim Interface As Object
    UserVar("BENEFICIARIO") =""
    Set Interface =CreateBennerObject("CA043.Autorizacao")
    Interface.Executar(CurrentSystem,0,CurrentQuery.FieldByName("AUTORIZACAO").AsInteger,0)
    Set Interface =Nothing
  Else
    BsShowMessage("Não existe autorização vinculada a esta cotação", "I")
  End If
End Sub



Public Sub TABLE_AfterScroll()
  If VisibleMode Then
    Dim DLL As Object
    Dim vFornecedor As Long
    Dim sql As Object
    Set sql = NewQuery

    vFornecedor  = LocalizaFornecedor(CurrentUser)

    Set DLL = CreateBennerObject("BSPORTALWEB.WEB")
    DLL.SelecionarFornecedor(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,vFornecedor)

    sql.Clear
    sql.Add("SELECT COUNT(1) QTD")
    sql.Add("  FROM SAM_LICITACAO_FORNECEDOR B")
    sql.Add(" WHERE LICITACAO = :LICITACAO")
    sql.Add("   AND FORNECEDOR = :FORNECEDOR")
    sql.ParamByName("FORNECEDOR").AsInteger = vFornecedor
    sql.ParamByName("LICITACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.Active = True
    If sql.FieldByName("QTD").AsInteger = 0 Then
      InfoDescription = "Fornecedor não credenciado a participar desta cotação. Não fornece produto cotado."
    Else
      DLL.PrepararItensFornecedor(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,vFornecedor)
    End If
    Set DLL = Nothing
    Set sql = Nothing
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("DATAHORAABERTURA").AsDateTime < ServerNow Then
    CanContinue = False
    BsShowMessage("Data de início deve ser maior que a data e hora atual", "E")
  End If

  CurrentQuery.FieldByName("DATAENCERRAMENTO").AsDateTime = DateAdd("h",CurrentQuery.FieldByName("PRAZO").AsInteger,CurrentQuery.FieldByName("DATAHORAABERTURA").AsDateTime)

End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "PARTICIPARLICITACAO" Then
    On Error GoTo Erro
    Dim sql As Object
    Dim vContador As Long
    Dim vAno As Integer
    Dim vDataEntrega As Date
    Dim vFornecedor As Long
    Dim vMensagem As String

    Set sql = NewQuery

    vFornecedor = LocalizaFornecedor(CurrentUser)
    vMensagem = ""
    vMensagem = VerificaLanceMaiorPrecoMaximo(vFornecedor,CurrentQuery.FieldByName("HANDLE").AsInteger)

    If Len(vMensagem) = 0 Then

      vDataEntrega = ServerNow

      vAno = Year(vDataEntrega)
      NewCounter("WEB_LICITACAO",vAno,1,vContador)


      If Not InTransaction Then
        StartTransaction
      End If

      sql.Clear
      sql.Add("UPDATE SAM_LICITACAO_FORNECEDOR SET DATAHORAENTREGA = :DATA, PROTOCOLO = :PROTOCOLO")
      sql.Add("WHERE FORNECEDOR = :FORNECEDOR AND LICITACAO = :LICITACAO")
      sql.ParamByName("DATA").AsDateTime = vDataEntrega
      sql.ParamByName("PROTOCOLO").AsInteger = Str(vAno)+Format(vContador,"0000")
      sql.ParamByName("FORNECEDOR").AsInteger = LocalizaFornecedor(CurrentUser)
      sql.ParamByName("LICITACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      sql.ExecSQL

      If InTransaction Then
        Commit
      End If

      InfoDescription = "Por favor, anote o número do protocolo: " + Str(vAno)+Format(vContador,"0000")
      Set sql = Nothing
      Exit Sub
    Else
      'Caso exista algum valor maior que o máximo nenhum dos produtos será considerado.
      'Desta forma, apagar o campo protocolo e este será desconsiderado da contabilização do
      'vencedor.
      If Not InTransaction Then
        StartTransaction
      End If

      sql.Clear
      sql.Add("UPDATE SAM_LICITACAO_FORNECEDOR SET DATAHORAENTREGA = :DATA, PROTOCOLO = NULL")
      sql.Add("WHERE FORNECEDOR = :FORNECEDOR AND LICITACAO = :LICITACAO")
      sql.ParamByName("DATA").AsDateTime = ServerNow
      sql.ParamByName("FORNECEDOR").AsInteger = LocalizaFornecedor(CurrentUser)
      sql.ParamByName("LICITACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      sql.ExecSQL

      If InTransaction Then
        Commit
      End If


      CanContinue = False
      CancelDescription = vMensagem + "Caso os valores não sejam alterados, o lance será desconsiderado."
      Set sql = Nothing
      Exit Sub
    End If

    Erro:
      CanContinue = False
      CancelDescription = Err.Description
      Set sql = Nothing
      If InTransaction Then
        Rollback
      End If
  ElseIf CommandID = "BOTAOFECHARCOTACAO" Then
    BOTAOFECHARCOTACAO_OnClick
  End If
End Sub
