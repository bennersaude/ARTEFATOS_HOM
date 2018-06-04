'HASH: 6977ACE0BB0F2966F2F7AF8D0B19E112
Option Explicit

Public Function CHECAR_PARAM_ROTINACARTAO

  CHECAR_PARAM_ROTINACARTAO = True

  Dim SQLP As Object
  Set SQLP = NewQuery
  SQLP.Add("SELECT DESTINOCARTAOAVULSO FROM SAM_PARAMETROSBENEFICIARIO")
  SQLP.Active = True

  If SQLP.EOF Then
    CHECAR_PARAM_ROTINACARTAO = False
    MsgBox("Ploblemas com parametros Rotina sem parâmetros.")
    Exit Function
  End If

  If SQLP.FieldByName("DESTINOCARTAOAVULSO").Value = "I" Then
    Exit Function
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT B.NUMERO, C.TABTIPOGERACAO, C.ARQUIVOCONTRATO, C.ARQUIVO, C.TABTIPOLEIAUTE,")
  SQL.Add("       C.TABEXPORTACAO, C.ARQUIVOFAMILIA, C.ARQUIVOBENEFICIARIO                   ")
  SQL.Add("  FROM CA_ROTSOLICIT A,                                                           ")
  SQL.Add("       CA_ROTSOLICITPARAM B,                                                      ")
  SQL.Add("       SAM_ROTINACARTAO C                                                         ")
  SQL.Add(" WHERE A.HANDLE = :HROTSOLICIT                                                    ")
  SQL.Add("   AND B.ROTSOLICIT = A.HANDLE                                                    ")
  SQL.Add("   AND C.ROTSOLICITPARAM = B.HANDLE                                               ")
  SQL.ParamByName("HROTSOLICIT").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    'CHECAR_PARAM_ROTINACARTAO =False
    'MsgBox("Rotina sem parâmetros.")
    Exit Function
  End If

  Dim vErro As Boolean
  vErro = False

  Do While Not SQL.EOF

    If SQL.FieldByName("TABTIPOLEIAUTE").AsInteger = 1 Then
      If SQL.FieldByName("TABEXPORTACAO").AsInteger = 1 Then
        If SQL.FieldByName("ARQUIVOCONTRATO").AsString = "" _
                            Or SQL.FieldByName("ARQUIVOFAMILIA").AsString = "" _
                            Or SQL.FieldByName("ARQUIVOBENEFICIARIO").AsString = "" Then
          MsgBox("Os parâmetros da rotina estão incompletos verifique os campos 'Nome do arquivo por contrato', '...por familia' e '...por beneficiário'!")
          CHECAR_PARAM_ROTINACARTAO = False
          vErro = True
          Exit Do
        End If
      Else
        If SQL.FieldByName("ARQUIVO").AsString = "" Then
          MsgBox("Os parâmetros da rotina estão incompletos verifique o campo 'Arquivo'!")
          CHECAR_PARAM_ROTINACARTAO = False
          vErro = True
          Exit Do
        End If
      End If
    End If

    SQL.Next

  Loop

  If vErro Then
    CHECAR_PARAM_ROTINACARTAO = False
    Exit Function
  End If

  CHECAR_PARAM_ROTINACARTAO = True

End Function


Public Sub BOTAOCANCELAR_OnClick()
  If(CurrentQuery.State <>1)Then
  MsgBox("Processo em edição.")
  Exit Sub
End If

If CurrentQuery.FieldByName("SITUACAO").Value = "A" Then
  MsgBox("Rotina não gerada.")
  Exit Sub
End If

If CurrentQuery.FieldByName("SITUACAO").Value = "P" Then
  MsgBox("Rotina já processada. Cancelamento não permitido.")
  Exit Sub
End If


Dim retorno As Boolean
Dim INTERFACE As Object
'Set interface=CreateBennerObject("CA019.Processos")

'retorno=interface.Cancela_Geral(CurrentQuery.FieldByName("handle").Value)'Handle da rotina de solicitacao
Set interface = CreateBennerObject("CA020.Cancelar")

retorno = interface.CancRotSolicit(CurrentSystem, CurrentQuery.FieldByName("handle").Value)'Handle da rotina de solicitacao

If retorno = False Then
  MsgBox("Processo não concluído.")
  Exit Sub
End If

RefreshNodesWithTable("CA_ROTSOLICIT")


End Sub

Public Sub BOTAOGERAR_OnClick()
  If(CurrentQuery.State <>1)Then
  MsgBox("Processo em edição.")
  Exit Sub
End If

If CurrentQuery.FieldByName("SITUACAO").Value = "G" Then
  MsgBox("Rotina já gerada")
  Exit Sub
End If

Dim SQL As Object
Set SQL = NewQuery
SQL.Add("SELECT HANDLE FROM CA_ROTSOLICITPARAM WHERE ROTSOLICIT = :HANDLE")
SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True

If SQL.EOF Then
  MsgBox("Rotina sem parâmetros.")
  Exit Sub
End If

Set SQL = Nothing

Set SQL = NewQuery
SQL.Add("SELECT HANDLE FROM CA_ROTSOLICITPARAM WHERE ROTSOLICIT = :HANDLE AND SITUACAO = 'A'")
SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True

If SQL.EOF Then
  MsgBox("Rotina Já processada.")
  Exit Sub
End If

Dim retorno As Boolean
Dim numero As Integer

Dim INTERFACE As Object
Set INTERFACE = CreateBennerObject("CA019.Processos")
numero = RecordHandleOfTable("SAM_FILIAL")
If numero >= 0 Then
  retorno = INTERFACE.Gerar(CurrentSystem, CurrentQuery.FieldByName("handle").Value)
Else
  retorno = INTERFACE.GerarTodasFiliais(CurrentSystem, CurrentQuery.FieldByName("handle").Value)
End If

If retorno = False Then
  MsgBox("Processo não concluído.")
  Exit Sub
End If

RefreshNodesWithTable("CA_ROTSOLICIT")


End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If(CurrentQuery.State <>1)Then
  MsgBox("Processo em edição.")
  Exit Sub
End If

If CurrentQuery.FieldByName("SITUACAO").Value = "P" Then
  MsgBox("Rotina já processada")
  Exit Sub
End If

Dim SQL As Object

Set SQL = NewQuery
SQL.Add("SELECT HANDLE FROM CA_ROTSOLICITPARAM WHERE ROTSOLICIT = :HANDLE AND SITUACAO = 'G'")
SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True

If SQL.EOF Then
  MsgBox("Rotina não gerada.")
  Exit Sub
End If

If CHECAR_PARAM_ROTINACARTAO = False Then
  Exit Sub
End If

Dim retorno As Boolean
Dim INTERFACE As Object
Set INTERFACE = CreateBennerObject("CA019.PROCESSOS")

retorno = INTERFACE.Processar(CurrentSystem, CurrentQuery.FieldByName("handle").Value)

If retorno = False Then
  MsgBox("Processo não concluído.")
End If

RefreshNodesWithTable("CA_ROTSOLICIT")


End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  If CurrentQuery.FieldByName("SITUACAO").Value <>"A" Then
    MsgBox("Operação não permitida.")
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DESCRICAO").AsString = "" Then
    MsgBox("Descrição obrigatória.")
    CANCONTINUE = False
    DESCRICAO.SetFocus
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABPROCESSO").AsInteger = 2 Then
    If CurrentQuery.FieldByName("SERVICOS").AsString = "" Then
      MsgBox("Serviço obrigatório.")
      CANCONTINUE = False
      Exit Sub
    End If
  End If
End Sub

Public Sub TABLE_NewRecord()


  Dim prFilial As Long
  Dim prFilialProcessamento As Long
  Dim prMsg As String

  BuscarFiliais(CurrentSystem, prFilial, prFilialProcessamento, prMsg)
  CurrentQuery.FieldByName("filial").AsInteger = prFilial
  CurrentQuery.FieldByName("filialprocessamento").AsInteger = prFilialProcessamento

End Sub

