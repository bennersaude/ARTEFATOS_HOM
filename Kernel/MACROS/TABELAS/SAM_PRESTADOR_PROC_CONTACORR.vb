'HASH: CF6ADB12094D846EA4CAA2773EAF8E33
'#Uses "*bsShowMessage"

Option Explicit

Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Dim S As Object
  Set S = NewQuery
  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  'If S.FieldByName("CONTROLEDEACESSO").Value = "N" Then
  '  Ok = True
  '  Set S=Nothing
  '  Exit Function
  'End If

  'SQL.Add("SELECT DATAFINAL,RESPONSAVEL FROM SAM_PRESTADOR_PROC WHERE HANDLE = :HANDLE")
  'SQL.ParamByName("HANDLE").Value=RecordHandleOfTable("SAM_PRESTADOR_PROC")
  'SQL.Active=True
  'Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser,True,False)

  SQL.Add("SELECT SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.FILIALPADRAO FROM SAM_PRESTADOR_PROC, SAM_PRESTADOR WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PROC.PRESTADOR")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And ((SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser) Or (SQL.FieldByName("FILIALPADRAO").IsNull)), True, False)

  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida." + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Function VerificaDataFinal
  If Not Ok Then
    bsShowMessage(Mensagem, "E")
    BANCO.ReadOnly = True
    AGENCIA.ReadOnly = True
    CONTACORRENTE.ReadOnly = True
    DV.ReadOnly = True
  Else
    BANCO.ReadOnly = False
    AGENCIA.ReadOnly = False
    CONTACORRENTE.ReadOnly = False
    DV.ReadOnly = False
  End If
End Function

Public Sub TABLE_AfterInsert()
  'Dim PRESTADOR As Object
  'Set PRESTADOR = NewQuery
  '  PRESTADOR.Clear
  '  PRESTADOR.Add("SELECT DISTINCT C.CCNOME,              ")
  '  PRESTADOR.Add("       C.CCCPF,                        ")
  '  PRESTADOR.Add("       C.NAOGERARDOCUMENTO,            ")
  '  PRESTADOR.Add("       C.NAOCOBRARTARIFA               ")
  '  PRESTADOR.Add("  FROM SAM_PRESTADOR_PROC_CONTACORR PC,")
  '  PRESTADOR.Add("       SAM_PRESTADOR_PROC_CREDEN PD,   ")
  '  PRESTADOR.Add("       SAM_PRESTADOR_PROC P,           ")
  '  PRESTADOR.Add("       SFN_CONTAFIN C                  ")
  '  PRESTADOR.Add(" WHERE PD.PRESTADORPROCESSO = P.HANDLE ")
  '  PRESTADOR.Add("   And P.PRESTADOR = C.PRESTADOR       ")
  '  PRESTADOR.Add("   And PD.HANDLE = :HANDLE             ")
  '  PRESTADOR.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
  '  PRESTADOR.Active = True
  '
  '  CurrentQuery.FieldByName("CCNOME").Value = PRESTADOR.FieldByName("CCNOME").Value
  '  CurrentQuery.FieldByName("CCCPF").Value = PRESTADOR.FieldByName("CCCPF").Value
  '  CurrentQuery.FieldByName("NAOGERARDOCUMENTO").Value = PRESTADOR.FieldByName("NAOGERARDOCUMENTO").Value
  '  CurrentQuery.FieldByName("NAOCOBRARTARIFA").Value = PRESTADOR.FieldByName("NAOCOBRARTARIFA").Value

  'Set PRESTADOR = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  VerificaDataFinal
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If
  VerificaDataFinal
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If
  VerificaDataFinal
End Sub

'#uses "*CheckCPFCNPJ"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Msg As String
  If Not CurrentQuery.FieldByName("CCCPFCNPJ").IsNull Then
    If Not CheckCPFCNPJ(CurrentQuery.FieldByName("CCCPFCNPJ").AsString, 0, True, Msg) Then
      bsShowMessage(Msg, "E")
      CanContinue = False
    End If
  End If


  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT BANCO, AGENCIA, CONTACORRENTE, DV ")
  SQL.Add("  FROM SFN_CONTAFIN WHERE PRESTADOR = :PRESTADOR")
  SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  SQL.Active = True

  If (SQL.FieldByName("BANCO").Value = CurrentQuery.FieldByName("BANCO").Value) Then
    If (SQL.FieldByName("AGENCIA").Value = CurrentQuery.FieldByName("AGENCIA").Value) Then
      If (SQL.FieldByName("CONTACORRENTE").Value = CurrentQuery.FieldByName("CONTACORRENTE").Value) Then
        If (SQL.FieldByName("DV").Value = CurrentQuery.FieldByName("DV").Value) Then
          bsShowMessage("A conta igual a conta vigente nao pode ser alterada!", "E")
          Set SQL = Nothing
          CanContinue = False
          Exit Sub
        End If
      End If
    End If
  End If

  Set SQL = Nothing

  Dim vPermitidas(31) As String
  Dim I As Integer
  Dim vDV1 As String
  Dim vDV2 As String

  vPermitidas(0) = "+"
  vPermitidas(1) = "-"
  vPermitidas(2) = "*"
  vPermitidas(3) = "/"
  vPermitidas(4) = "."
  vPermitidas(5) = ","
  vPermitidas(6) = ";"
  vPermitidas(7) = ":"
  vPermitidas(8) = "'"
  vPermitidas(9) = "="
  vPermitidas(10) = "|"
  vPermitidas(11) = "_"
  vPermitidas(12) = ")"
  vPermitidas(13) = "("
  vPermitidas(14) = "%"
  vPermitidas(15) = "$"
  vPermitidas(16) = "#"
  vPermitidas(17) = "@"
  vPermitidas(18) = "!"
  vPermitidas(19) = "?"
  vPermitidas(20) = "~"
  vPermitidas(21) = "`"
  vPermitidas(22) = """"
  vPermitidas(23) = "{"
  vPermitidas(24) = "}"
  vPermitidas(25) = "["
  vPermitidas(26) = "]"
  vPermitidas(27) = "^"
  vPermitidas(28) = "\"
  vPermitidas(29) = ">"
  vPermitidas(30) = "<"

  vDV1 = Mid(CurrentQuery.FieldByName("DV").AsString, 1, 1)
  vDV2 = Mid(CurrentQuery.FieldByName("DV").AsString, 2, 1)

  For I = 0 To 30
    If (vDV1 = vPermitidas(I)) Or (vDV2 = vPermitidas(I)) Then
      bsShowMessage("Dígito verificador da conta corrente não pode conter caracteres especiais!", "E")
      CanContinue = False
      Exit Sub
    End If
  Next I

  Dim interface As Object
  Dim vBanco As Long
  Dim vAgencia As Long
  Dim vConta As String
  Dim vDV As String
  Set interface = CreateBennerObject("FINANCEIRO.CONTAFIN")

  vBanco = CurrentQuery.FieldByName("BANCO").AsInteger
  vAgencia = CurrentQuery.FieldByName("AGENCIA").AsInteger
  vConta = CurrentQuery.FieldByName("CONTACORRENTE").AsString
  vDV = CurrentQuery.FieldByName("DV").AsString

  Dim vsMensagem As String

  If Not interface.VerificaDV(CurrentSystem, vBanco, vAgencia, CurrentQuery.TQuery, vConta, vDV, vsMensagem) Then
    bsShowMessage(vsMensagem, "E")
    CanContinue = False
    Set interface = Nothing
    Exit Sub
  End If

  Set interface = Nothing
End Sub

Public Sub AGENCIA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_AGENCIA.NOME|SFN_AGENCIA.AGENCIA"
  vCriterio = "SFN_AGENCIA.BANCO =" + CurrentQuery.FieldByName("BANCO").AsString
  vCampos = "Nome|Código"

  vHandle = interface.Exec(CurrentSystem, "SFN_AGENCIA", vColunas, 1, vCampos, vCriterio, "Tabela de Agências", True, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("AGENCIA").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub BANCO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "NOME|CODIGO"

  vCampos = "Nome|Código"

  vHandle = interface.Exec(CurrentSystem, "SFN_BANCO", vColunas, 1, vCampos, vCriterio, "Tabela de Bancos", True, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BANCO").Value = vHandle
    CurrentQuery.FieldByName("AGENCIA").Clear
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_NewRecord()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT TABGERACAOREC FROM SFN_CONTAFIN WHERE PRESTADOR = :PRESTADOR")
  SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  SQL.Active = True

  If Not SQL.EOF And SQL.FieldByName("TABGERACAOREC").AsInteger <> 1 Then
    bsShowMessage("A conta financeira do prestador não é do tipo conta corrente !", "E")
    Set SQL = Nothing
    RefreshNodesWithTable("SAM_PRESTADOR_PROC_CREDEN")
    Exit Sub
  End If
  Set SQL = Nothing
End Sub
