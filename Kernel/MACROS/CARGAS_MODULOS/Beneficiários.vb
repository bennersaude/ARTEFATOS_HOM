'HASH: 0C74893346CC536394FBD61E073EE5DF
'#Uses "*bsShowMessage"
'BUILDER

Public Sub BENCONSULTACC_OnClick()
Dim interface As Object

Set interface =CreateBennerObject("SFNContaCorrente.Consulta")
    interface.Executar(CurrentSystem)
Set interface =Nothing
End Sub

Public Sub CONSULTABENEFICIARIO_OnClick()
Dim interface As Object

'Set interface =CreateBennerObject("CA010.ConsultaBeneficiario")
'Alterado SMS 90338 - Rodrigo Andrade 30/11/2007 -
'Separação da Interface da regra de negocio para consulta de Beneficiários
Set interface =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
interface.filtro(CurrentSystem,1,"")

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
      qPermissao.Add("    or EXISTS(SELECT 1")
      qPermissao.Add("                FROM CLI_RECURSO")
      qPermissao.Add("               WHERE PRESTADOR = RU.PRESTADOR))")
      qPermissao.ParamByName("USUARIO").AsInteger =CurrentUser
      qPermissao.Active =True
      If qPermissao.EOF Then
        bsShowMessage("Usuário sem permissão para acessar o prontuário médico!","I")
        Exit Sub
      End If

      qTabela.Clear
      qTabela.Add("SELECT NOME")
      qTabela.Add("  FROM Z_TABELAS ")
      qTabela.Add(" WHERE HANDLE = :HANDLE")
      qTabela.ParamByName("HANDLE").AsInteger = CurrentTable
      qTabela.Active =True
      If qTabela.FieldByName("NOME").AsString ="SAM_MATRICULA" Then
        vMatricula =CurrentQuery.FieldByName("HANDLE").AsInteger
      Else
        If RecordHandleOfTable("SAM_MATRICULA")>0 Then
          vMatricula =RecordHandleOfTable("SAM_MATRICULA")
        Else
          bsShowMessage("É necessário selecionar um cadastro na pasta ''Matrícula Única'' do módulo Beneficiários!","I")
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

Public Sub DIGITABENEFICIARIO_OnClick()
  Dim dllBSInterface001 As Object

  Set dllBSInterface001 = CreateBennerObject("BSINTERFACE0011.DigitarBeneficiario")

  dllBSInterface001.Exec(CurrentSystem, _
						 0, _
						 0, _
						 0)

  Set dllBSInterface001 = Nothing
End Sub

Public Sub GERASOLICITAUXPAS_OnClick()
  'Balani SMS 50888 19/10/2005
  Dim Interface As Object
  Set Interface = CreateBennerObject("SAMSOLICITAUX.ROTINAS")
  Interface.GeraSolicitacao(CurrentSystem)
  Set Interface = Nothing
  bsShowMessage("Processo concluído!", "I")
  'final SMS 50888
End Sub


Public Sub MIGRABENEFLOTE_OnClick()

  Dim Interface As Object
  Dim mensagem As String
  Set Interface = CreateBennerObject("CONTRATO.BENEFICIARIO")

  On Error GoTo erro
    mensagem = Interface.MigrarLote(CurrentSystem)

    Set Interface = Nothing
    Exit Sub

  Erro:
    Set Interface = Nothing

    bsShowMessage(Err.Description, "I")
End Sub

Public Sub ROTINAENCERRASUSPENS_OnClick()
Dim Interface As Object

Set Interface =CreateBennerObject("BsBen014.EncerramentoSuspensao")
    Interface.Processar(CurrentSystem)
Set Interface =Nothing
End Sub
