'HASH: 3A264DB2E2E9765B0886DDF5A89EF58A
'MACRO SAM_CONTRATO_CENTROCUSTO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub CENTROCUSTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CENTROCUSTO.ESTRUTURA|SFN_CENTROCUSTO.DESCRICAO|SFN_CENTROCUSTO.CODIGOREDUZIDO"

  vCriterio = "HANDLE>0 AND ULTIMONIVEL = 'S' "

  vCampos = "Estrutura|Descrição|Código"

  vHandle = interface.Exec(CurrentSystem, "SFN_CENTROCUSTO", vColunas, 1, vCampos, vCriterio, "Centro de Custo", False, CENTROCUSTO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CENTROCUSTO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()

 If WebMode Then
  	CONTRATO.ReadOnly = True
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT TABCENTROCUSTO FROM SAM_CONTRATO WHERE HANDLE=:HCONTRATO")
  REGIAO.Visible = False 'sms 49081 jacinto
  SQL.ParamByName("HCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TABCENTROCUSTO").AsInteger = 3 Then 'CASO NO CONTRATO FOR LOTAÇÃO MOSTRAR O CAMPO "LOTAÇÃO"
    LOTACAO.Visible = True
  Else
    LOTACAO.Visible = False
    If SQL.FieldByName("TABCENTROCUSTO").AsInteger = 2 Then 'sms 49081 jacinto
       REGIAO.Visible = True 'sms 49081 jacinto
    End If
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q As Object
  Dim qBusca As Object
  Dim vHandleContCC As Integer
  Dim viTabCentroCusto As Integer

  Set q = NewQuery
  Set qBusca = NewQuery

  If CurrentQuery.FieldByName("FILIAL").IsNull Then
	bsShowMessage("O Campo Filial deve ser preenchido", "E")
    CanContinue = False
  End If

  q.Add("SELECT ULTIMONIVEL FROM SFN_CENTROCUSTO WHERE HANDLE=:HANDLE")
  q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CENTROCUSTO").AsInteger
  q.Active = True
  If q.FieldByName("ULTIMONIVEL").AsString <>"S" Then
    bsShowMessage("Centro de custo deve ser de último nível", "E")
    CanContinue = False
  End If

  q.Clear
  q.Active = False
  q.Add("SELECT TABCENTROCUSTO FROM SAM_CONTRATO WHERE HANDLE=:HCONTRATO")
  q.ParamByName("HCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  q.Active = True

  viTabCentroCusto = q.FieldByName("TABCENTROCUSTO").AsInteger

  If q.FieldByName("TABCENTROCUSTO").AsInteger <>3 Then 'CASO NO CONTRATO NÃO FOR LOTAÇÃO
    CurrentQuery.FieldByName("LOTACAO").Clear
  End If

  'Coelho - SMS: 34294
  'CASO NO CONTRATO FOR FILIAL:
  'NÃO UTILIZAR O MESMO CENTRO DE CUSTO NA MESMA FILIAL
'  If q.FieldByName("TABCENTROCUSTO").AsInteger = 2 Then
'    qBusca.Clear
'    qBusca.Active = False
'    qBusca.Add("SELECT HANDLE FROM SAM_CONTRATO_CENTROCUSTO")
'    qBusca.Add("WHERE CONTRATO =:HCONTRATO")
'    qBusca.Add("AND FILIAL =:HFILIAL")
'    qBusca.ParamByName("HCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
'    qBusca.ParamByName("HFILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger
'    qBusca.Active = True

'    While Not qBusca.EOF
'      vHandleContCC = qBusca.FieldByName("HANDLE").AsInteger
'      If (vHandleContCC > 0) And (vHandleContCC <> CurrentQuery.FieldByName("HANDLE").AsInteger) Then
'        MsgBox("Já existe Centro de custo cadastrado para esta filial!")
'        CanContinue = False
'        Exit Sub
'      End If
'      qBusca.Next
'    Wend
'  End If

  'CASO NO CONTRATO FOR LOTAÇÃO:
  'NÃO UTILIZAR O MESMO CENTRO DE CUSTO NA MESMA FILIAL COM A MESMA LOTAÇÃO
'  If q.FieldByName("TABCENTROCUSTO").AsInteger = 3 Then
'    qBusca.Clear
'    qBusca.Active = False
'    qBusca.Add("SELECT HANDLE FROM SAM_CONTRATO_CENTROCUSTO")
'    qBusca.Add("WHERE CONTRATO =:HCONTRATO")
'    qBusca.Add("AND FILIAL =:HFILIAL")
'    If CurrentQuery.FieldByName("LOTACAO").AsInteger <> 0 Then
'      qBusca.Add("AND LOTACAO =:HLOTACAO")
'    Else
'      qBusca.Add("AND LOTACAO IS NULL")
'    End If
'    qBusca.ParamByName("HCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
'    qBusca.ParamByName("HFILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger
'    If CurrentQuery.FieldByName("LOTACAO").AsInteger <> 0 Then
'      qBusca.ParamByName("HLOTACAO").AsInteger = CurrentQuery.FieldByName("LOTACAO").AsInteger
'    End If
'    qBusca.Active = True
'
'    While Not qBusca.EOF
'      vHandleContCC = qBusca.FieldByName("HANDLE").AsInteger
'      If (vHandleContCC > 0) And (vHandleContCC <> CurrentQuery.FieldByName("HANDLE").AsInteger) Then
'        MsgBox("Lotação já cadastrada para este contrato nesta filial!")
'        CanContinue = False
'        Exit Sub
'      End If
'      qBusca.Next
'    Wend
'  End If

  'sms 49081 - Jacinto
  q.Clear
  q.Active = False
  q.Add("SELECT COUNT(1) QTDE")
  q.Add("  FROM SAM_CONTRATO_CENTROCUSTO")
  q.Add(" WHERE CONTRATO    = :CONTRATO")
  q.Add("   AND FILIAL      = :FILIAL")
  q.Add("   And CENTROCUSTO = :CENTROCUSTO")
  If (viTabCentroCusto = 3) And Not (CurrentQuery.FieldByName("LOTACAO").IsNull) Then
    q.Add("   And LOTACAO     = :LOTACAO")
  Else
    If (viTabCentroCusto = 3) And (CurrentQuery.FieldByName("LOTACAO").IsNull) Then
      q.Add("   And LOTACAO     IS NULL")
    End If
  End If
  If (viTabCentroCusto = 2) And Not (CurrentQuery.FieldByName("REGIAO").IsNull) Then
    q.Add("   And REGIAO      = :REGIAO")
  Else
     If (viTabCentroCusto = 2) And (CurrentQuery.FieldByName("REGIAO").IsNull) Then
       q.Add("   And REGIAO      IS NULL")
     End If
  End If
  If CurrentQuery.FieldByName("CONTRATO").IsNull Then
    q.ParamByName("CONTRATO").DataType = ftInteger
    q.ParamByName("CONTRATO").Clear
  Else
    q.ParamByName("CONTRATO").AsInteger    = CurrentQuery.FieldByName("CONTRATO").AsInteger
  End If
  If CurrentQuery.FieldByName("FILIAL").IsNull Then
    q.ParamByName("FILIAL").DataType = ftInteger
    q.ParamByName("FILIAL").Clear
  Else
    q.ParamByName("FILIAL").AsInteger      = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If
  If CurrentQuery.FieldByName("CENTROCUSTO").IsNull Then
    q.ParamByName("CENTROCUSTO").DataType = ftInteger
    q.ParamByName("CENTROCUSTO").Clear
  Else
    q.ParamByName("CENTROCUSTO").AsInteger = CurrentQuery.FieldByName("CENTROCUSTO").AsInteger
  End If
  If viTabCentroCusto = 3 Then
    If Not CurrentQuery.FieldByName("LOTACAO").IsNull Then
      q.ParamByName("LOTACAO").AsInteger     = CurrentQuery.FieldByName("LOTACAO").AsInteger
    End If
  End If
  If viTabCentroCusto = 2 Then
    If Not CurrentQuery.FieldByName("REGIAO").IsNull Then
      q.ParamByName("REGIAO").AsInteger      = CurrentQuery.FieldByName("REGIAO").AsInteger
    End If
  End If
  q.Active = True

  If q.FieldByName("QTDE").AsInteger > 0 Then
     bsShowMessage("Já existe uma configuração idêntica para este centro de custo!", "E")
     CanContinue = False
     Exit Sub
  End If
  'fim sms 49081


  Set q = Nothing
  Set qBusca = Nothing

End Sub

