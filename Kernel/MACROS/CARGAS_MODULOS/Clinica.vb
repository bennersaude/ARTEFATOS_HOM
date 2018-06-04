'HASH: 2A3BF5A5A8FA2991E33A106BC8F8A7A7
 

Public Sub AGENDA_OnClick()
  Dim AGENDA As Object

  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.Agendamento(CurrentSystem, 0, 0, 0, 0, 0, 0)

  Set AGENDA = Nothing

End Sub

Public Sub ATENDAMBULATORIAL_OnClick()
  Dim BSCLI004 As Object
  Set BSCLI004 = CreateBennerObject("BSCLI004.Ambulatorial")
  BSCLI004.AtendimentoAmbulatorial(CurrentSystem)
  Set BSCLI004 = Nothing
End Sub

Public Sub CONSULTAAGENDAMENTO_OnClick()
  Dim AGENDA As Object
  Set AGENDA =CreateBennerObject("BSCli002.Rotinas")
  AGENDA.Agendamento(CurrentSystem)
  Set AGENDA =Nothing
End Sub

Public Sub CONSULTAPRECOMATMED_OnClick()
  Dim Interface As Object
  Set Interface =CreateBennerObject("BSCLI015.PEGAPRECO")
  Interface.PRECOMATMED(CurrentSystem)
  Set Interface =Nothing
End Sub

Public Sub DIGITACAOAUTORIZACAO_OnClick()
      Dim Interface As Object
      'autorizacao antiga
      'Set INTERFACE =CreateBennerObject("ca014.Autorizacao")
      'INTERFACE.Exec(CurrentSystem,0,0,0)
      'autorizacao nova
      Set Interface = CreateBennerObject("CA043.Autorizacao")
      Interface.Executar(CurrentSystem,0, 0, 0)
      Set Interface =Nothing
End Sub

Public Sub CONSULTAPRONTUARIO_OnClick()
  Dim vMatricula As Long
  Dim qPermissao As Object
  Dim qTabela As Object
  Set qTabela =NewQuery
  Set qPermissao =NewQuery

  qPermissao.Clear
  qPermissao.Add("SELECT 1")
  qPermissao.Add("  FROM CLI_RECURSO_USUARIO RU")
  qPermissao.Add(" WHERE RU.USUARIO = :USUARIO")
  qPermissao.Add("   AND (RU.ACESSAATENDIMENTOS = 'S'")
  qPermissao.Add("    OR EXISTS(SELECT 1")
  qPermissao.Add("                FROM CLI_RECURSO")
  qPermissao.Add("               WHERE PRESTADOR = RU.PRESTADOR))")
  qPermissao.ParamByName("USUARIO").AsInteger =CurrentUser
  qPermissao.Active =True
  If qPermissao.EOF Then
    MsgBox "Usuário sem permissão para acessar o prontuário médico!"
    Exit Sub
  End If

  qTabela.Clear
  qTabela.Add("SELECT NOME")
  qTabela.Add("  FROM Z_TABELAS ")
  qTabela.Add(" WHERE HANDLE = :HANDLE")
  qTabela.ParamByName("HANDLE").AsInteger =CurrentTable
  qTabela.Active =True
  If qTabela.FieldByName("NOME").AsString ="SAM_MATRICULA" Then
    If CurrentQuery.Active Then
      vMatricula =CurrentQuery.FieldByName("HANDLE").AsInteger
    Else
      If RecordHandleOfTable("SAM_MATRICULA")>0 Then
        vMatricula =RecordHandleOfTable("SAM_MATRICULA")
      Else
        MsgBox "É necessário selecionar um paciente na pasta ''Matrícula'' do módulo Clínica!"
        Exit Sub
      End If
    End If
  Else
    If RecordHandleOfTable("SAM_MATRICULA")>0 Then
      vMatricula =RecordHandleOfTable("SAM_MATRICULA")
    Else
      MsgBox "É necessário selecionar um paciente na pasta ''Matrícula'' do módulo Clínica!"
      Exit Sub
    End If
  End If
  qTabela.Active =False

  Dim ATE As Object
  Set ATE =CreateBennerObject("BSCLI004.ROTINAS")
  Dim SQL As Object
  Set SQL =NewQuery
  SQL.Active =False
  SQL.Clear
  SQL.Add("SELECT MAX(HANDLE) BEN")
  SQL.Add("  FROM SAM_BENEFICIARIO")
  SQL.Add(" WHERE MATRICULA = :MATRICULA")
  SQL.Add("   AND DATACANCELAMENTO IS NULL")
  SQL.ParamByName("MATRICULA").AsInteger =vMatricula
  SQL.Active =True
  If Not SQL.FieldByName("BEN").IsNull Then 'EXISTE UM BENEFICIÁRIO VÁLIDO
    ATE.Atendimento(CurrentSystem,0,0,SQL.FieldByName("BEN").AsInteger)
  Else
    ATE.Atendimento(CurrentSystem,0,vMatricula,0)
  End If

  Set qTabela =Nothing
  Set SQL =Nothing
  Set ATE =Nothing
  Set qPermissao =Nothing
End Sub
