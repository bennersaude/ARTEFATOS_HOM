'HASH: 24EF87329E92B5B8BFFCD312B47F0EE7
'#Uses "*bsShowMessage"
'#uses "*QueryToXML"

Option Explicit
Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not AtendimentoEncerrado(CurrentQuery.FieldByName("SUBJETIVO").AsInteger) Then
    If Not CurrentQuery.FieldByName("RESPOSTA").IsNull Then
      CanContinue = False
      bsShowMessage("Não é possível responder o pedido, o atendimento ainda está em aberto!","E")
    End If
  End If
  If Not AtendimentoSessaoEncerrado Then
    If Not CurrentQuery.FieldByName("RESPOSTA").IsNull Then
      CanContinue = False
      bsShowMessage("Não é possível responder o pedido, o atendimento selecionado está encerrado!","E")
    End If
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim HSubjetivo As Long

  HSubjetivo = CLng(SessionVar("CLI_SUBJETIVO"))

  If (HSubjetivo > 0) Then
	If AtendimentoEncerrado(CLng(SessionVar("CLI_SUBJETIVO"))) Then
      CanContinue = False
      bsShowMessage("Não é possível realizar um encaminhamento, o atendimento selecionado está encerrado!","E")
	End If
  Else
    CanContinue = False
    bsShowMessage("É necessário selecionar um atendimento antes de realizar um encaminhamento!","E")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.InInsertion Then
    CurrentQuery.FieldByName("SUBJETIVO").Value = SessionVar("CLI_SUBJETIVO")
    CurrentQuery.FieldByName("DATAENVIO").Value = Now
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

  If CommandID = "ENCAMINHAR" Then
    Encaminhar
    bsShowMessage("Pedido encaminhado com sucesso!","I")
  End If

  If CommandID = "FINALIZARPEDIDO" Then
  	FinalizarPedido
	bsShowMessage("Encaminhamento finalizado com sucesso!","I")
  End If

  If CommandID = "VALIDARATENDIMENTO" Then
  	If Not ValidarAtendimento Then
	  CanContinue = False
	  bsShowMessage("O atendimento deste encaminhamento já está encerrado, não é possível editar!","E")
	End If
  End If

  If CommandID = "PROBLEMAPRINCIPAL" Then
	If IncluirProblemaPrincipal Then
	  bsShowMessage("Problema principal incluído com sucesso!","I")
	Else
	  bsShowMessage("Problema principal não encontrado para o paciente!","I")
	End If
  End If

  If CommandID = "BUSCARPROBLEMAS" Then
    InfoDescription = BuscarProblemas
  End If

  If CommandID = "BUSCARPROBLEMASSELECIONADOS" Then
    InfoDescription = BuscarProblemasSelecionados
  End If

End Sub

Function ValidarAtendimento As Boolean

  Dim query As Object

  Set query = NewQuery

  query.Clear
  query.Add("SELECT A.DATAENCERRAMENTO  ")
  query.Add("  FROM CLI_SUBJETIVO      A                           ")
  query.Add("  JOIN CLI_ENCAMINHAMENTO E ON E.SUBJETIVO = A.HANDLE ")
  query.Add(" WHERE E.SUBJETIVO = :SUBJETIVO")
  query.ParamByName("SUBJETIVO").AsInteger = CurrentQuery.FieldByName("SUBJETIVO").AsInteger
  query.Active = True

  If query.FieldByName("DATAENCERRAMENTO").IsNull Then
    ValidarAtendimento = True
  Else
    ValidarAtendimento = False
  End If

  Set query = Nothing

End Function
Function Encaminhar As Boolean

  Dim query As Object

  Set query = NewQuery

  query.Clear
  query.Add("UPDATE CLI_ENCAMINHAMENTO")
  query.Add("   SET DATAENVIO = :DATAENVIO")
  query.Add(" WHERE HANDLE = :HANDLE")
  query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  query.ParamByName("DATAENVIO").AsDateTime = Now
  query.ExecSQL

  Set query = Nothing

End Function

Function IncluirProblemaPrincipal As Boolean
  Dim query As Object
  Dim sql As Object

  Set sql = NewQuery

  sql.Clear
  sql.Add(" SELECT D.HANDLE                                               ")
  sql.Add("   FROM CLI_PACIENTEDIAGNOSTICO D                              ")
  sql.Add("   JOIN CLI_USUARIO_SESSAO      U ON U.MATRICULA = D.MATRICULA ")
  sql.Add("  WHERE D.EHCIDPRINCIPAL = 'S'                                 ")
  sql.Add("    AND U.USUARIO = :USUARIO                                   ")
  sql.ParamByName("USUARIO").AsInteger = CurrentUser
  sql.Active = True

  If sql.EOF Then
    IncluirProblemaPrincipal =False
  Else

    Set query = NewQuery

    query.Clear
    query.Add("INSERT INTO CLI_ENCAMINHAMENTO_PROBLEMA          ")
    query.Add("        (HANDLE, ENCAMINHAMENTO, DIAGNOSTICO)    ")
    query.Add("VALUES (:HANDLE, :ENCAMINHAMENTO, :DIAGNOSTICO)  ")
    query.ParamByName("HANDLE").AsInteger = NewHandle("CLI_ENCAMINHAMENTO_PROBLEMA")
    query.ParamByName("ENCAMINHAMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    query.ParamByName("DIAGNOSTICO").AsInteger = sql.FieldByName("HANDLE").AsInteger
    query.ExecSQL

    IncluirProblemaPrincipal = True

    Set query = Nothing

  End If

  Set sql = Nothing

End Function
Function AtendimentoEncerrado(handle As Long) As Boolean
    Dim sql As Object

    Set sql = NewQuery
	sql.Clear
	sql.Add("SELECT DATAENCERRAMENTO      ")
	sql.Add("  FROM CLI_SUBJETIVO         ")
	sql.Add(" WHERE HANDLE = :HANDLE")
	sql.ParamByName("HANDLE").Value = handle
	sql.Active = True

    If Not sql.FieldByName("DATAENCERRAMENTO").IsNull Or sql.EOF Then
      AtendimentoEncerrado = True
    Else
      AtendimentoEncerrado = False
    End If

    Set sql = Nothing
End Function
Function AtendimentoSessaoEncerrado As Boolean
    Dim sql As Object

    Set sql = NewQuery
	sql.Clear
	sql.Add("SELECT S.DATAENCERRAMENTO                               ")
	sql.Add("  FROM CLI_SUBJETIVO S                                  ")
	sql.Add("  JOIN CLI_USUARIO_SESSAO U ON U.ATENDIMENTO = S.HANDLE ")
	sql.Add(" WHERE U.USUARIO = :USUARIO                             ")
	sql.ParamByName("USUARIO").AsInteger = CurrentUser
	sql.Active = True

    If Not sql.FieldByName("DATAENCERRAMENTO").IsNull Or sql.EOF Then
      AtendimentoSessaoEncerrado = True
    Else
      AtendimentoSessaoEncerrado = False
    End If

    Set sql = Nothing
End Function
Public Function BuscarProblemas As String
  Dim sql As BPesquisa

  Set sql = NewQuery

  sql.Clear
  sql.Add(" SELECT D.HANDLE, D.DATACONSULTA, C2.ESTRUTURA, C2.DESCRICAO, D.EHCIDPRINCIPAL  ")
  sql.Add("   FROM CLI_PACIENTEDIAGNOSTICO   D                             ")
  sql.Add("   JOIN CLI_CID_CID               C1 ON C1.HANDLE = D.CID       ")
  sql.Add("   JOIN SAM_CID                   C2 ON C2.HANDLE = C1.CID      ")
  sql.Add("  ORDER BY D.DATACONSULTA DESC       ")
  sql.Active = True

  BuscarProblemas = QueryToXml(sql)

  Set sql = Nothing

End Function
Public Function BuscarProblemasSelecionados As String
  Dim sql As BPesquisa

  Set sql = NewQuery

  sql.Clear
  sql.Add(" SELECT D.HANDLE, C2.ESTRUTURA, C2.DESCRICAO, D.EHCIDPRINCIPAL     ")
  sql.Add("   FROM CLI_ENCAMINHAMENTO_PROBLEMA P                              ")
  sql.Add("   JOIN CLI_PACIENTEDIAGNOSTICO     D  ON D.HANDLE = P.DIAGNOSTICO ")
  sql.Add("   JOIN CLI_CID_CID                 C1 ON C1.HANDLE = D.CID        ")
  sql.Add("   JOIN SAM_CID                     C2 ON C2.HANDLE = C1.CID       ")
  sql.Add("  WHERE P.ENCAMINHAMENTO = :HANDLE                                 ")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  BuscarProblemasSelecionados = QueryToXml(sql)

  Set sql = Nothing

End Function
Public Function FinalizarPedido As String
  Dim query As Object

  Set query = NewQuery

  query.Clear
  query.Add("UPDATE CLI_ENCAMINHAMENTO")
  query.Add("   SET DATAFINALIZACAO = :DATAFINALIZACAO")
  query.Add(" WHERE HANDLE = :HANDLE")
  query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  query.ParamByName("DATAFINALIZACAO").AsDateTime = Now
  query.ExecSQL

  Set query = Nothing

End Function
