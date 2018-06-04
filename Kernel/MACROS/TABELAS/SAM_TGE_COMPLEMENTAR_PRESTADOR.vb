'HASH: FE5D9A9CFB16107EAB334082E269F8B1
'Macro: SAM_TGE_COMPLEMENTAR_PRESTADOR
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"



Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub EVENTOAGERAR_OnExit()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT CIRURGICO FROM SAM_TGE WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  SQL.Active = True
  If SQL.FieldByName("CIRURGICO").AsString = "S" Then
    CODIGOPAGTO.ReadOnly = True
  Else
    CODIGOPAGTO.ReadOnly = False
  End If
  Set SQL = Nothing
End Sub

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

  vColunas = "GRAU|DESCRICAO"

  If CurrentQuery.FieldByName("EVENTOAGERAR").IsNull Then
    vCriterio = "HANDLE = -1"
  Else
    vCriterio = "HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTOAGERAR").AsString + ")"
  End If

  vCampos = "Grau|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterio, "Tabela De Graus", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAUAGERAR").Value = vHandle
  End If
  Set interface = Nothing
End Sub



Public Sub TABLE_AfterScroll()
  CLASSEASSOCIADO.Visible = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim TGE As Object


  ' Na TGE não existe mais o grau principal
  '  Set TGE=NewQuery
  '  TGE.Add("SELECT HANDLE,GRAU FROM SAM_TGE WHERE HANDLE = :HANDLE")
  '  TGE.ParamByName("HANDLE").Value=RecordHandleOfTable("SAM_TGE")
  '  TGE.Active=True
  '  MesmoEvento=TGE.FieldByName("HANDLE").AsInteger =CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  '  MesmoGrau  =TGE.FieldByName("GRAU").AsInteger   =CurrentQuery.FieldByName("GRAUAGERAR").AsInteger
  '  If MesmoGrau And MesmoEvento Then '
  '	 CanContinue =False
  '     MsgBox "Mesmo Evento/Grau. Operação não Permitida."
  '     Set TGE=Nothing
  '	 Exit Sub
  '  End If

  Set SQL = NewQuery
  SQL.Add("SELECT COUNT(*) T")
  SQL.Add("  FROM SAM_TGE_COMPLEMENTAR_PRESTADOR")
  SQL.Add(" WHERE EVENTOAGERAR = :EVENTOAGERAR")
  SQL.Add("   AND GRAUAGERAR = :GRAUAGERAR")
  SQL.Add("   AND HANDLE <> :HEVENTOCOMPLEMENTAR")
  SQL.Add("   AND PRESTADOR = :PRESTADOR")
  SQL.ParamByName("HEVENTOCOMPLEMENTAR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("EVENTOAGERAR").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  SQL.ParamByName("GRAUAGERAR").Value = CurrentQuery.FieldByName("GRAUAGERAR").AsInteger
  SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
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
    bsShowMessage(Msg, "E")
  End If
End If
End If

Set SQL = Nothing
End Sub

Public Sub EVENTOAGERAR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOAGERAR.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOAGERAR").Value = vHandle
  End If
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


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
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
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

