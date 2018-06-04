'HASH: 426CCFF32B077C98E211F6221ED80F19
'Macro: SAM_TGE_COMPLEMENTAR_REDE

'Última alteração: Milton/17/01/2002 -SMS 5976

'#Uses "*bsShowMessage"

Public Sub GRAUAGERAR_OnChange()
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT COUNT (*) REC FROM SAM_TGE_GRAU WHERE EVENTO = :EVENTO AND GRAU = :GRAU")
  Q.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  Q.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAUAGERAR").AsInteger
  Q.Active = True
  If Q.FieldByName("REC").AsInteger = 0 Then
    CurrentQuery.FieldByName("GRAUAGERAR").Clear
  End If
End Sub

Public Sub GRAUAGERAR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "GRAU|DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  If CurrentQuery.FieldByName("EVENTOAGERAR").IsNull Then
    vCriterio = "HANDLE = -1"
  Else
    vCriterio = "HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTOAGERAR").AsString + ")"
  End If

  vCampos = "Grau|Descrição|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterio, "Tabela De Graus", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAUAGERAR").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()

  If (WebMode) Then
    GRAUAGERAR.WebLocalWhere = "A.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTOAGERAR))"
    EVENTOAGERAR.WebLocalWhere = "A.ULTIMONIVEL = 'S' "
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim TGE As Object
  Dim msg As String

  Set SQL = NewQuery
  SQL.Add("SELECT COUNT(*) T FROM SAM_TGE_COMPLEMENTAR_REDE WHERE EVENTOAGERAR = :EVENTOAGERAR And GRAUAGERAR = :GRAUAGERAR AND REDERESTRITA = :REDERESTRITA AND HANDLE <> :HANDLE")
  SQL.ParamByName("EVENTOAGERAR").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  SQL.ParamByName("GRAUAGERAR").Value = CurrentQuery.FieldByName("GRAUAGERAR").AsInteger
  SQL.ParamByName("REDERESTRITA").Value = CurrentQuery.FieldByName("REDERESTRITA").AsInteger
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If SQL.FieldByName("T").AsInteger >0 Then
    CanContinue = False
    bsShowMessage("Registro Duplicado! Operação não permitida.", "E")
  End If
  SQL.Active = False

  SQL.Clear
  SQL.Add("SELECT A.CALCCODPAGTOEVENTOCIRURGICO A,")
  SQL.Add("       B.CIRURGICO B")
  SQL.Add("  FROM SAM_PARAMETROSATENDIMENTO A,")
  SQL.Add("       SAM_TGE B")
  SQL.Add(" WHERE B.HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  SQL.Active = True

  'Se o evento é cirúrgico não pode ser informado o codigo do pagamento
  If(SQL.FieldByName("A").AsString = "S")Then
  If(SQL.FieldByName("B").AsString = "S")Then
  If Not(CurrentQuery.FieldByName("CODIGOPAGTO").IsNull)Then
    CanContinue = False
    msg = "Está marcado nos parâmetros gerais que o percentual de pagamento" + Chr(13)
    msg = msg + "será calculado pelo sistema para eventos cirúrgicos." + Chr(13)
    msg = msg + "O campo Código de pagamento deverá ser deixado em branco"
    bsShowMessage(msg, "E")
  End If
End If
End If

SQL.Active = False
Set SQL = Nothing

End Sub

Public Sub EVENTOAGERAR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOAGERAR").Value = vHandle
    CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  End If
  Set interface = Nothing

End Sub


Public Sub CODIGOPAGTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PERCENTUALPGTO.CODIGOPAGTO|DESCRICAO|INCIDENCIAMINIMA|PERCENTUALPGTOINCIDENCIA1|PERCENTUALPGTODEMAIS|USADOAUTORIZACAO|USADOPAGAMENTO"

  vCampos = "Código|Descrição|Incidência Mínima|% Pagto Inc 1|% Pagto Demais|Usado Autorização|Usado Pagto"

  vHandle = interface.Exec(CurrentSystem, "SAM_PERCENTUALPGTO", vColunas, 1, vCampos, vCriterio, "Tabela de Códigos de Pagamentos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CODIGOPAGTO").Value = vHandle
  End If

  Set interface = Nothing

End Sub

