'HASH: C0144607C820677D7C7E98D309D09C70
'Macro: SAM_PRESTADOR_PROC_MEMBROS

'#Uses "*bsShowMessage"

'Mauricio Ibelli - 04/01/2002 - sms3165 - Se filial padrao do prestador for nulo não checar responsavel

Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Mensagem = ""

  Dim S As Object
  Set S = NewQuery

  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  'Garcia
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
    Mensagem = "Processo finalizado!" + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

'Public Sub xMEMBRO_OnPopup(ShowPopup As Boolean)
'
'	UpdateLastUpdate("SAM_PRESTADOR")
'	If CurrentQuery.FieldByName("OPERACAO").AsString = "E" Then
'		MEMBRO.LocalWhere = " HANDLE IN (SELECT PRESTADOR FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE ENTIDADE = "+CurrentQuery.FieldByName("PRESTADOR").AsString+")"
'	Else
'		MEMBRO.LocalWhere = ""
'	End If
'End Sub

Public Sub MEMBRO_OnPopup(ShowPopup As Boolean)

	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vFiltro As String

	Dim qAux As Object
	Set qAux = NewQuery

	Dim interface As Object
	Set interface = CreateBennerObject("Procura.Procurar")

	Dim qProc As BPesquisa
	Set qProc = NewQuery
	qProc.Add("SELECT P.PRESTADOR                                                      ")
	qProc.Add("  FROM SAM_PRESTADOR_PROC P                                             ")
	qProc.Add("  JOIN SAM_PRESTADOR_PROC_CREDEN C ON (P.HANDLE = C.PRESTADORPROCESSO)  ")
	qProc.Add(" WHERE C.HANDLE = :CREDENCIAMENTO                                       ")

	qProc.ParamByName("CREDENCIAMENTO").AsInteger  = RecordHandleOfTable("SAM_PRESTADOR_PROC_CREDEN")
	qProc.Active = True

	CurrentQuery.FieldByName("PRESTADOR").AsInteger = qProc.FieldByName("PRESTADOR").AsInteger


	qAux.Add("SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE = " + qProc.FieldByName("PRESTADOR").AsString)
	qAux.Active = True

	ShowPopup = False

	Dim Msg As String

	vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.Z_NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
	vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"


	If qAux.FieldByName("FISICAJURIDICA").AsInteger = 2 Then
		vCriterio = "SAM_PRESTADOR.HANDLE <> " + qProc.FieldByName("PRESTADOR").AsString + vFiltro
	Else
		vCriterio = "SAM_PRESTADOR.HANDLE = " + qProc.FieldByName("PRESTADOR").AsString
	End If

	vCampos = "Código|Nome do Prestador|Data Cred.|Categoria|Estados|Município"

	vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|*ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|*MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 1, vCampos, vCriterio, "Prestador", True, MEMBRO.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("MEMBRO").Value = vHandle
	End If

	ShowPopup = False
	Set qProc = Nothing
	Set interface = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_MEMBROS"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q As Object
  Set Q = NewQuery
  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_MEMBROS"
    Exit Sub
  End If

  'Incluído na SMS 59739 - 29.03.2006 - não permitir cadastrar o mesmo membro
  Q.Clear
  Q.Add("SELECT 1 ")
  Q.Add("  FROM SAM_PRESTADOR_PROC_MEMBROS             ")
  Q.Add(" WHERE PRESTADORPROCESSO = :PRESTADORPROCESSO ")
  Q.Add("   AND PRESTADOR = :PRESTADOR                 ")
  Q.Add("   AND MEMBRO = :MEMBRO                       ")
  Q.Add("   AND HANDLE <> :HANDLE                      ")
  Q.ParamByName("PRESTADORPROCESSO").AsInteger = CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
  Q.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  Q.ParamByName("MEMBRO").AsInteger = CurrentQuery.FieldByName("MEMBRO").AsInteger
  Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.Active = True
  If (Not Q.EOF) Then
    bsShowMessage("Membro já cadastrado para este processo.", "E")
    CanContinue = False
    Set Q = Nothing
    Exit Sub
  End If
  Set Q = Nothing
  'Final SMS 59739

  Dim SQL As Object
  Set SQL = NewQuery
  'BY WILSON
  'VERIFICAR SE O PRESTADOR É PESSOA JURIDICA
  SQL.Add("SELECT A.FISICAJURIDICA FROM SAM_PRESTADOR A WHERE A.HANDLE = :PREST")
  SQL.ParamByName("PREST").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.Active = True
  If SQL.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
    CanContinue = False
    bsShowMessage("Operação não permitida para prestador - Pessoa Física!", "E")
    Exit Sub
  End If
  'END BY WILSON

  SQL.Clear
  SQL.Add("SELECT * FROM SAM_PRESTADOR_PRESTADORDAENTID A")
  SQL.Add(" WHERE A.ENTIDADE = :PREST")
  SQL.Add("   AND A.PRESTADOR = :MEMBRO")
  SQL.Add("   AND A.DATAINICIAL <= :DATA")
  SQL.Add("   AND (:DATA <= A.DATAFINAL OR A.DATAFINAL IS NULL)")
  SQL.ParamByName("PREST").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.ParamByName("MEMBRO").Value = CurrentQuery.FieldByName("MEMBRO").AsInteger

  If CurrentQuery.FieldByName("OPERACAO").AsString = "E" Then
    SQL.ParamByName("DATA").Value = ServerDate()
  Else
    SQL.ParamByName("DATA").Value = CurrentQuery.FieldByName("datainicial").AsDateTime
  End If

  SQL.Active = True

  If CurrentQuery.FieldByName("OPERACAO").AsString = "E" Then
    If SQL.EOF Then
      CanContinue = False
      bsShowMessage("Membro não encontrado. Operação incoerente", "E")
      Exit Sub
    End If
  End If
  '*****************************************************
  'Incio da Alteração
  'Alterado por Durval em 22/04/2002
  'Alterado por Garcia em 08/04/2002

  If Not SQL.EOF Then
    If (Not SQL.FieldByName("DATAFINAL").IsNull) And (SQL.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) Then
      bsShowMessage("Data Final inferior a data de cadastramento !", "E")
      CanContinue = False
    End If
  End If
  'Fim da alteração
  '*****************************************************
  SQL.Active = False
  Set SQL = Nothing

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
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("datainicial").AsDateTime = ServerDate
  CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser
End Sub
