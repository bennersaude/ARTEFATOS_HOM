'HASH: 0E13252E5984390BC80B8228D7612BF7
'macro: SAM_AUDITORIA

'#Uses "*bsShowMessage"
'Última alteração: Milton/17/01/2002 -SMS 5976





Public Sub AUDITOR_OnPopup(ShowPopup As Boolean)

  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "IDENTIFICACAO|NOME"


  vCriterio = "FILIALAUDITOR = " + CurrentQuery.FieldByName("FILIALDESTINO").AsString

  vCampos = "Identificação|Nome"

  vHandle = interface.Exec(CurrentSystem, "SAM_AUDITOR", vColunas, 1, vCampos, vCriterio, "Auditor", True, "")

  CurrentQuery.Edit
  CurrentQuery.FieldByName("AUDITOR").Value = vHandle

  ShowPopup = False
End Sub

Public Sub CIDDIAGNOSTICADO_OnPopup(ShowPopup As Boolean)
  CIDDIAGNOSTICADO.AnyLevel = True
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|Z_DESCRICAO"

  vCriterio = ""

  vCampos = "Estrutura|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_CID", vColunas, 1, vCampos, vCriterio, "CID", True, "")

  CurrentQuery.Edit
  CurrentQuery.FieldByName("CIDDIAGNOSTICADO").Value = vHandle

  ShowPopup = False



End Sub


Public Sub TABLE_AfterScroll()
  'UpdateLastUpdate("SAM_AUDITOR")
  If WebMode Then
	  AUDITOR.WebLocalWhere = "FUNCAO = 'R' OR FUNCAO = 'A' AND FILIALAUDITOR = @CAMPO(FILIALDESTINO)"
  ElseIf VisibleMode Then
	  AUDITOR.LocalWhere = "FUNCAO = 'R' OR FUNCAO = 'A' AND FILIALAUDITOR = @FILIALDESTINO"
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Exclusão de auditoria permitida apenas pela interface de Autorização", "E")
  End If
End Sub




Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Alteração de auditoria permitida apenas pela interface de Autorização", "E")
  End If
  Dim SQL As String
  SQL = "SELECT FILIALAUDITOR FROM SAM_AUDITOR WHERE FUNCAO <> 'P'"

  If WebMode Then
  	FILIALDESTINO.WebLocalWhere = "A.HANDLE IN (" + SQL + ")"
  ElseIf VisibleMode Then
	FILIALDESTINO.WebLocalWhere = "FILIAIS.HANDLE IN (" + SQL + ")"
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Inclusão de auditoria permitida apenas pela interface de Autorização", "E")
  End If

  Dim SQL As String
  SQL = "SELECT FILIALAUDITOR FROM SAM_AUDITOR WHERE FUNCAO <> 'P'"

  If WebMode Then
  	FILIALDESTINO.WebLocalWhere = "A.HANDLE IN (" + SQL + ")"
  ElseIf VisibleMode Then
	FILIALDESTINO.WebLocalWhere = "FILIAIS.HANDLE IN (" + SQL + ")"
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  If CurrentQuery.FieldByName("TABREGRAAUDITAR").AsInteger = 2 Then
    CurrentQuery.FieldByName("MOTIVOGLOSA").Clear
  End If

  If Not CurrentQuery.FieldByName("PARECER").IsNull Then
    If CurrentQuery.FieldByName("PARECERDATA").IsNull Then
      bsShowMessage("Data do parecer é obrigatório.", "I")
      ConContinue = False
      Exit Sub
    End If
  End If

  If Not CurrentQuery.FieldByName("PARECERDATA").IsNull Then
    If CurrentQuery.FieldByName("PARECER").IsNull Then
      bsShowMessage("Parecer é obrigatório.", "I")
      ConContinue = False
      Exit Sub
    End If
  End If

  'verifica data de pedido
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT DATAAUTORIZACAO FROM SAM_AUTORIZ WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
  SQL.Active = True
  If ((Format(CurrentQuery.FieldByName("DATAPEDIDO").AsDateTime, "DD/MM/YYYY") < Format(SQL.FieldByName("DATAAUTORIZACAO").AsDateTime, "DD/MM/YYYY"))Or _
      (Format(CurrentQuery.FieldByName("DATAPEDIDO").AsDateTime, "DD/MM/YYYY") > ServerDate))Then
    bsShowMessage("Data de pedido anterior a data da autorização ou é uma data futura", "I")
    ConContinue = False
  Exit Sub
End If
Set SQL = Nothing
End Sub

Public Sub TABLE_NewRecord()
  Dim SQL As Object
  Set SQL = NewQuery
  Dim qAudit As Object
  Set qAudit = NewQuery


  SQL.Clear
  SQL.Add("SELECT B.FILIALCUSTO FROM SAM_AUTORIZ A, SAM_BENEFICIARIO B WHERE A.HANDLE = :HANDLE AND B.HANDLE = A.BENEFICIARIO")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_AUTORIZ")
  SQL.Active = True

  If (SQL.FieldByName("FILIALCUSTO").AsInteger > 0) Then
    CurrentQuery.FieldByName("FILIALORIGEM").Value = SQL.FieldByName("FILIALCUSTO").AsInteger
    CurrentQuery.FieldByName("FILIALDESTINO").Value = SQL.FieldByName("FILIALCUSTO").AsInteger
  End If

  Set SQL = Nothing

  CurrentQuery.FieldByName("DATAPEDIDO").Value = ServerDate

  Dim SQL2 As Object
  Set SQL2 = NewQuery
  SQL2.Clear
  SQL2.Add("SELECT MOTIVOAUDITORIAALERTA FROM SAM_PARAMETROSATENDIMENTO")
  SQL2.Active = True

  CurrentQuery.FieldByName("MOTIVOAUDITORIA").Value = SQL2.FieldByName("MOTIVOAUDITORIAALERTA").AsInteger

  Set SQL2 = Nothing


  qAudit.Clear
  qAudit.Add("SELECT AU.HANDLE HANDLE_AUDITOR, AU.NOME NOME_AUDITOR FROM")
  qAudit.Add("                 SAM_AUDITOR AU,")
  qAudit.Add("                 SAM_PRESTADOR PA,")
  qAudit.Add("                 SAM_PRESTADOR PRECEB,")
  qAudit.Add("                 SAM_PRESTADOR_RELACIONADO PR,")
  qAudit.Add("                 SAM_PARAMETROSPRESTADOR PARAM,")
  qAudit.Add("                 SAM_AUTORIZ A,")
  qAudit.Add("                 SAM_AUTORIZ_EVENTOSOLICIT ES")
  qAudit.Add("WHERE PA.HANDLE = AU.PRESTADOR")
  qAudit.Add("AND A.HANDLE = ES.AUTORIZACAO")
  qAudit.Add("AND ES.RECEBEDOR = PRECEB.HANDLE")
  qAudit.Add("AND PA.HANDLE = PR.PRESTADORRELACIONADO")
  qAudit.Add("AND PRECEB.HANDLE = PR.PRESTADOR")
  qAudit.Add("AND PR.RELACAO = PARAM.RELACAOAUDITOR")
  qAudit.Add("AND A.HANDLE=:HANDLE")
  qAudit.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_AUTORIZ")
  qAudit.Active = True
  If qAudit.FieldByName("HANDLE_AUDITOR").AsInteger > 0 Then
    CurrentQuery.FieldByName("AUDITOR").AsInteger = qAudit.FieldByName("HANDLE_AUDITOR").AsInteger
  End If

  Set qAudit = Nothing


End Sub

