'HASH: AEBC00F93D6C91F1D612C48A67D2ACA0
'Macro: SAM_FAMILIA
'Juliano -24/11/2000 -Atribuição Do titular e da região
'#Uses "*bsShowMessage"
Option Explicit

Dim viDiaCobrancaAnterior           As Integer
Dim viTabResponsavelAnterior        As Integer
Dim viHTitularResponsavelAnterior   As Long
Dim vsCobrancaDeEventoAnterior      As String
Dim vsNumeroBenefAutomaticoAnterior As String
Dim vsModoEdicao                    As String

Public Sub BOTAOALTERARADESAO_OnClick()
  Dim vcContainer     As Object
  Dim BSINTERFACE0002 As Object
  Dim vsMensagem      As String
  Dim viRetorno       As Long

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
								   1, _
								   "TV_FORM0086", _
								   "Alterar Data Adesão", _
								   0, _
								   300, _
								   310, _
								   False, _
								   vsMensagem, _
								   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
	Case -1
		bsShowMessage("Operação cancelada pelo usuário!", "I")
	Case 1
		bsShowMessage(vsMensagem, "I")
  End Select
  Set BSINTERFACE0002 = Nothing
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("A Família já está cancelada !", "I")
    Exit Sub
  End If

  Dim vsMensagemErro As String
  Dim viRetorno As Integer
  Dim vvContainer As CSDContainer
  Dim Interface As Object

  Set vvContainer = NewContainer

  SessionVar("HFAMILIA_CANCELAMENTO") = CurrentQuery.FieldByName("HANDLE").AsString

  Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")

  viRetorno = Interface.Exec(CurrentSystem, _
                             1, _
                             "TV_FORM0016", _
                             "Cancelamento de Família", _
                             0, _
                             180, _
                             420, _
                             False, _
                             vsMensagemErro, _
                             vvContainer)

  Set vvContainer = Nothing
  Set Interface = Nothing

  If viRetorno =  1 Then
      bsShowMessage(vsMensagemErro, "I")
  End If
End Sub

Public Sub BOTAOCONSULTAVALORMODULO_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("A tabela não pode estar em edição", "E")
    Exit Sub
  End If

  Dim Interface As Object

  Set Interface = CreateBennerObject("SAMConsultaBenef.Consultas")
  Interface.Executar(CurrentSystem, 2, 0, CurrentQuery.FieldByName("HANDLE").AsInteger, 0)
  Set Interface = Nothing
End Sub

Public Sub BOTAOCRIAPESSOARESPONSAVEL_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
  Else
    Dim Interface As Object
    Dim vsMensagem As String

    Set Interface = CreateBennerObject("BSBEN020.Familia")
    Interface.CriarPessoa(CurrentSystem, _
                          CurrentQuery.FieldByName("HANDLE").AsInteger, _
                          vsMensagem)

    bsShowMessage(vsMensagem, "I")

    Set Interface = Nothing

    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAODECLARACAO_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If

  Dim Interface As Object

  Set Interface = CreateBennerObject("BSBEN011.Declaracao")
  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing

End Sub

Public Sub BOTAOFINANCEIRO_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  Dim viHContaFin As Long

  viHContaFin = RetornaContaFinanceira

  If viHContaFin > 0 Then
    Dim Interface   As Object
    Set Interface = CreateBennerObject("SAMCONTAFINANCEIRA.CONSULTA")
    Interface.Exec(CurrentSystem, viHContaFin)
    Set Interface = Nothing
  Else
    bsShowMessage("Conta financeira não encontrada", "I")
  End If

End Sub

Public Sub BOTAOGRIDBENEFICIARIOS_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If
  Dim OLEGrid As Object
  Dim vSQL As String

  vSQL = "SELECT NOME, BENEFICIARIO, DVCARTAO, MATRICULA, DATAADESAO, DATACANCELAMENTO FROM SAM_BENEFICIARIO WHERE FAMILIA = " + CurrentQuery.FieldByName("HANDLE").AsString + " ORDER BY NOME"

  Set OLEGrid = CreateBennerObject("SamGrid.DataSet")
  OLEGrid.Exec(CurrentSystem, vSQL, "Beneficiários da Família", "NOME")
  Set OLEGrid = Nothing
End Sub

Public Sub BOTAOINSCRICAO_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Add("SELECT LOCALFATURAMENTO FROM SAM_CONTRATO WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CONTRATO").Value
  Sql.Active = True

  If(Sql.FieldByName("LOCALFATURAMENTO").Value <>"F") Then
    Sql.Active = False
    Set Sql = Nothing
    bsShowMessage("Opção válida somente para contratos de faturamento na família", "I")
    Exit Sub
  End If
  Sql.Active = False
  Set Sql = Nothing

  Dim Interface As Object
  Set Interface = CreateBennerObject("SAMinscricao.Inscricao")
  Interface.Exec(CurrentSystem, "F", CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("CONTRATO").Value)
  Set Interface = Nothing
End Sub

Public Sub BOTAOMUDARRESPONSAVELFIN_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If
  Dim Interface As Object
  Dim vsMensagemErro As String
  Dim viRetorno As Integer
  Dim vvContainer As CSDContainer

  Set vvContainer = NewContainer

  SessionVar("HFAMILIA_TROCATITULAR") = CurrentQuery.FieldByName("HANDLE").AsString

  Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")

  viRetorno = Interface.Exec(CurrentSystem, _
                             1, _
                             "TV_FORM0002", _
                             "Mudança de Responsável Financeiro", _
                             0, _
                             310, _
                             310, _
                             False, _
                             vsMensagemErro, _
                             vvContainer)

  If viRetorno = 1 Then
    bsShowMessage(vsMensagemErro, "I")
  End If

  Set vvContainer = Nothing
  Set Interface = Nothing


End Sub

Public Sub BOTAOREATIVAR_OnClick()
  'No caso da macro estar sendo executada de dentro da interface de digitação
  'é possível que não exista nenhum registro selecionado
  If CurrentQuery.FieldByName("HANDLE").IsNull Then
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("A Família não está cancelada !", "I")
    Exit Sub
  End If

  'Se o titular estiver falecido,impedir reativar família
  Dim QryFalecimentoTitular As Object
  Set QryFalecimentoTitular = NewQuery
  QryFalecimentoTitular.Add("SELECT MOTIVOFALECIMENTOTITULAR, MOTIVOFALECIMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  QryFalecimentoTitular.Active = True
  If CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").Value = QryFalecimentoTitular.FieldByName("MOTIVOFALECIMENTOTITULAR").Value Then
    bsShowMessage("Necessário reativar tilular - Titular falecido !", "I")
    Exit Sub
  End If

  ' Se o CONTRATO estiver cancelado não pode
  Dim Sql As Object
  Set Sql = NewQuery
  Sql.Add("SELECT DATACANCELAMENTO FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  Sql.ParamByName("Contrato").Value = RecordHandleOfTable("SAM_CONTRATO")
  Sql.Active = True
  If Not Sql.FieldByName("DATACANCELAMENTO").IsNull Then
    Sql.Active = False
    Set Sql = Nothing
    bsShowMessage("Não é permitido reativar famílias nesse contrato - Contrato está Cancelado !", "I")
    Exit Sub
  End If
  Sql.Active = False
  Set Sql = Nothing

  Dim Interface As Object
  Dim vsMensagemErro As String
  Dim viRetorno As Integer
  Dim vvContainer As CSDContainer

  Set vvContainer = NewContainer

  SessionVar("HFAMILIA_REATIVACAO") = CurrentQuery.FieldByName("HANDLE").AsString

  Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")

  viRetorno = Interface.Exec(CurrentSystem, _
                             1, _
                             "TV_FORM0011", _
                             "Reativação de Família", _
                             0, _
                             120, _
                             230, _
                             False, _
                             vsMensagemErro, _
                             vvContainer)

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If

  If viRetorno = 1 Then
    bsShowMessage(vsMensagemErro, "I")
  End If

  Set vvContainer = Nothing
  Set Interface = Nothing

End Sub

Public Sub DATAADESAO_OnExit()

  If CurrentQuery.State = 3 Then
    CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").Value = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
  End If

End Sub

Public Sub DATAVENDA_OnChange()
  Dim SQLCONTRATO As Object
  Dim vDiaProjecao As Integer
  Dim vMes As Integer
  Dim vAno As Integer
  Set SQLCONTRATO = NewQuery
  SQLCONTRATO.Add("SELECT TABTIPOCONTRATO, GRUPOCONTRATO FROM SAM_CONTRATO WHERE HANDLE = :HCONTRATO")
  SQLCONTRATO.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  SQLCONTRATO.Active = True

  If SQLCONTRATO.FieldByName("TABTIPOCONTRATO").AsInteger <>1 Then
    Dim SQLPROJ As Object
    Set SQLPROJ = NewQuery
    SQLPROJ.Add("SELECT *")
    SQLPROJ.Add("FROM SAM_PROJECAOVENCIMENTO")
    SQLPROJ.Add("WHERE GRUPOCONTRATO = :GCONTRATO")
    SQLPROJ.ParamByName("GCONTRATO").Value = SQLCONTRATO.FieldByName("GRUPOCONTRATO").AsInteger
    SQLPROJ.Active = True
    SQLPROJ.First

    While Not SQLPROJ.EOF

      If(DatePart("d", CurrentQuery.FieldByName("DATAVENDA").AsDateTime)>= SQLPROJ.FieldByName("DIAVENDAINICIAL").AsInteger)And _
         (DatePart("d", CurrentQuery.FieldByName("DATAVENDA").AsDateTime)<= SQLPROJ.FieldByName("DIAVENDAFINAL").AsInteger)Then
      vDiaProjecao = SQLPROJ.FieldByName("DIAVENCIMENTO").AsInteger

      If SQLPROJ.FieldByName("QUALMES").AsInteger = 2 Then
        CurrentQuery.FieldByName("COBRANCANOMESSEGUINTE").Value = "S"
      End If

      vMes = Month(CurrentQuery.FieldByName("DATAVENDA").AsDateTime)
      vAno = Year(CurrentQuery.FieldByName("DATAVENDA").AsDateTime)

      If SQLPROJ.FieldByName("DIAVENCIMENTO").AsInteger <= DiasPorMes(vAno, vMes)Then
        vDiaProjecao = SQLPROJ.FieldByName("DIAVENCIMENTO").AsInteger
      Else
        vDiaProjecao = DiasPorMes(vAno, vMes)
      End If

      If vDiaProjecao <DatePart("d", CurrentQuery.FieldByName("DATAVENDA").AsDateTime)Then
        vMes = vMes + 1
        If vMes = 13 Then
          vMes = 1
          vAno = vAno + 1
        End If

        If vDiaProjecao >DiasPorMes(vAno, vMes)Then
          vDiaProjecao = DiasPorMes(vAno, vMes)
        End If
      End If

      CurrentQuery.FieldByName("DATAADESAO").AsDateTime = DateSerial(vAno, vMes, vDiaProjecao)
      CurrentQuery.FieldByName("DIACOBRANCA").Value = vDiaProjecao

    End If
    SQLPROJ.Next
  Wend
  Set SQLPROJ = Nothing

End If
Set SQLCONTRATO = Nothing

End Sub

Public Sub TABLE_AfterCancel()
  BOTAOCANCELAR.Visible              = True
  BOTAOFINANCEIRO.Visible            = True
  BOTAOREATIVAR.Visible              = True
  BOTAOCONSULTAVALORMODULO.Visible   = True
  BOTAOCRIAPESSOARESPONSAVEL.Visible = True
  BOTAODECLARACAO.Visible            = True
  BOTAOGRIDBENEFICIARIOS.Visible     = True
  BOTAOINSCRICAO.Visible             = True
  BOTAOMUDARRESPONSAVELFIN.Visible   = True

  If InTransaction Then
    Rollback
  End If
End Sub

Public Sub TABLE_AfterEdit()
  If CurrentQuery.FieldByName("REATIVAFAMCOMTITULAR").AsBoolean <> True Then
    CurrentQuery.FieldByName("REATIVAFAMCOMTITULAR").AsBoolean = False
  End If
End Sub

Public Sub TABLE_AfterInsert()
  Dim qContrato As Object
  Set qContrato = NewQuery

  qContrato.Clear
  qContrato.Add("SELECT LOCALFATURAMENTO, NUMEROFAMILIAAUTOMATICO")
  qContrato.Add("  FROM SAM_CONTRATO")
  qContrato.Add(" WHERE HANDLE = :CONTRATO")
  qContrato.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qContrato.Active = True

  If qContrato.FieldByName("LOCALFATURAMENTO").AsString = "F" Then
    PERIODOFATURAMENTO.ReadOnly = False
    DIACOBRANCA.ReadOnly        = False
  End If

  'Se família automática protege o campo
  If (qContrato.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString = "S") Then
    FAMILIA.ReadOnly= True
  Else
    FAMILIA.ReadOnly= False
  End If

  Set qContrato = Nothing

  Dim Interface   As Object

  Set Interface = CreateBennerObject("BSBEN020.Familia")

  Interface.NewRecord(CurrentSystem, _
                      CurrentQuery.TQuery)

  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterPost()
  Dim Interface   As Object
  Dim viResultado As Integer
  Dim vsMensagem  As String

  Set Interface = CreateBennerObject("BSBEN020.Familia")

  viResultado = Interface.AfterPost(CurrentSystem, _
                                    CurrentQuery.TQuery, _
                                    vsModoEdicao, _
                                    viTabResponsavelAnterior, _
                                    viHTitularResponsavelAnterior, _
                                    vsMensagem)

  Set Interface = Nothing

  If viResultado = 1 Then
    Err.Raise(vbsUserException, "", vsMensagem)
  Else
    If vsMensagem <> "" Then
      bsShowMessage(vsMensagem, "I")
    End If
  End If

  'Se estiver em modo desktop a transação deve ser iniciada
  'Isto é necessário devido ao fato dos componentes BVirtual não controlarem transação
  If CurrentQuery.IsVirtual And _
     VisibleMode And _
     InTransaction Then
    Commit
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.State <> 1 Then
    BOTAOCANCELAR.Visible              = False
    BOTAOFINANCEIRO.Visible            = False
    BOTAOREATIVAR.Visible              = False
    BOTAOCONSULTAVALORMODULO.Visible   = False
    BOTAOCRIAPESSOARESPONSAVEL.Visible = False
    BOTAODECLARACAO.Visible            = False
    BOTAOGRIDBENEFICIARIOS.Visible     = False
    BOTAOINSCRICAO.Visible             = False
    BOTAOMUDARRESPONSAVELFIN.Visible   = False

    If CurrentQuery.State = 2 Then
      PERIODOFATURAMENTO.ReadOnly = True
      FAMILIA.ReadOnly            = True
    End If
  Else
    BOTAOCANCELAR.Visible              = True
    BOTAOFINANCEIRO.Visible            = True
    BOTAOREATIVAR.Visible              = True
    BOTAOCONSULTAVALORMODULO.Visible   = True
    BOTAOCRIAPESSOARESPONSAVEL.Visible = True
    BOTAODECLARACAO.Visible            = True
    BOTAOGRIDBENEFICIARIOS.Visible     = True
    BOTAOINSCRICAO.Visible             = True
    BOTAOMUDARRESPONSAVELFIN.Visible   = True

    PERIODOFATURAMENTO.ReadOnly = True
    DIACOBRANCA.ReadOnly        = True
    FAMILIA.ReadOnly            = True
  End If


  If WebMode Then
    SessionVar("HCONTAFINANCEIRA_FAMILIA") = CStr(RetornaContaFinanceira)

    MOTIVOBLOQUEIO.WebLocalWhere = "A.HANDLE NOT IN (SELECT CONT.MOTIVOBLOQUEIOAUTOMATICO FROM SAM_CONTRATO CONT WHERE CONT.HANDLE = @CAMPO(CONTRATO) AND CONT.TABADESAORECEBIMENTO = 2)"
  Else
    MOTIVOBLOQUEIO.LocalWhere    = "SAM_MOTIVOBLOQUEIO.HANDLE NOT IN (SELECT CONT.MOTIVOBLOQUEIOAUTOMATICO FROM SAM_CONTRATO CONT WHERE CONT.HANDLE = @CONTRATO AND CONT.TABADESAORECEBIMENTO = 2)"
  End If

  SessionVar("HCONTRATO") = CStr(CurrentQuery.FieldByName("CONTRATO").AsInteger)

  PROXIMOVENCIMENTO.ReadOnly   = True
  DATAULTIMOREAJUSTE.ReadOnly  = True
  DIACOBRANCAORIGINAL.ReadOnly = True
  DATABLOQUEIO.ReadOnly        = True
  CONVENIO.ReadOnly            = True

  Dim Sql As Object
  Set Sql = NewQuery

  If (CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1) And _
     (Not CurrentQuery.FieldByName("TITULARRESPONSAVEL").IsNull) Then
    Sql.Clear
    Sql.Add("SELECT FAMILIA")
    Sql.Add("FROM SAM_BENEFICIARIO")
    Sql.Add("WHERE HANDLE = :HTITULARRESPONSAVEL")
    Sql.ParamByName("HTITULARRESPONSAVEL").AsInteger = CurrentQuery.FieldByName("TITULARRESPONSAVEL").AsInteger
    Sql.Active = True

    'Verificar se o responsável da família é um beneficiário da família atual
    If Sql.FieldByName("FAMILIA").AsInteger = _
       CurrentQuery.FieldByName("HANDLE").AsInteger Then
      Sql.Clear
      Sql.Add("SELECT LOCALFATURAMENTO")
      Sql.Add("FROM SAM_CONTRATO")
      Sql.Add("WHERE HANDLE = :HCONTRATO")
      Sql.ParamByName("HCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
      Sql.Active = True
      'Verificar se é faturamento na família
      If Sql.FieldByName("LOCALFATURAMENTO").AsString = "F" Then
        BOTAOMUDARRESPONSAVELFIN.Enabled = True
      Else
        BOTAOMUDARRESPONSAVELFIN.Enabled = False
      End If
    Else
      BOTAOMUDARRESPONSAVELFIN.Enabled = False
    End If
  Else
    BOTAOMUDARRESPONSAVELFIN.Enabled = False
  End If

  Sql.Clear
  If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
    Sql.Add("SELECT NOME FROM SAM_BENEFICIARIO")
    Sql.Add("WHERE HANDLE = :HBENEFICIARIO")
    Sql.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("TITULARRESPONSAVEL").AsInteger
    Sql.Active = True
    RESPONSAVEL.Text = "BENEFICIÁRIO: " + Sql.FieldByName("NOME").AsString
  Else
    Sql.Add("SELECT NOME FROM SFN_PESSOA")
    Sql.Add("WHERE HANDLE = :HPESSOA")
    Sql.ParamByName("HPESSOA").Value = CurrentQuery.FieldByName("PESSOARESPONSAVEL").AsInteger
    Sql.Active = True
    RESPONSAVEL.Text = "OUTRO: " + Sql.FieldByName("NOME").AsString
  End If

  Set Sql = Nothing

  'Verifica suspensão -Juliano 16-08-02----------------------------------------------------------------------------------------------

  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    CurrentQuery.FieldByName("CONTRATO").AsInteger, _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    vDataFinalSuspensao) Then
    BOTAOCANCELAR.Enabled = False
    BOTAOREATIVAR.Enabled = False
  Else
    BOTAOCANCELAR.Enabled = True
    BOTAOREATIVAR.Enabled = True
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Interface   As Object
  Dim viResultado As Integer
  Dim vsMensagem  As String

  Set Interface = CreateBennerObject("BSBEN020.Familia")

  viResultado = Interface.BeforeDelete(CurrentSystem, _
                                       CurrentQuery.TQuery, _
                                       vsMensagem)

  If viResultado = 1 Then
    CanContinue = False
    bsShowMessage(vsMensagem, "E")
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  BOTAOCANCELAR.Visible              = False
  BOTAOFINANCEIRO.Visible            = False
  BOTAOREATIVAR.Visible              = False
  BOTAOCONSULTAVALORMODULO.Visible   = False
  BOTAOCRIAPESSOARESPONSAVEL.Visible = False
  BOTAODECLARACAO.Visible            = False
  BOTAOGRIDBENEFICIARIOS.Visible     = False
  BOTAOINSCRICAO.Visible             = False
  BOTAOMUDARRESPONSAVELFIN.Visible   = False

  FAMILIA.ReadOnly            = True
  PERIODOFATURAMENTO.ReadOnly = True

  vsModoEdicao                    = "A"
  viDiaCobrancaAnterior           = CurrentQuery.FieldByName("DIACOBRANCA").AsInteger
  viTabResponsavelAnterior        = CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger
  vsCobrancaDeEventoAnterior      = CurrentQuery.FieldByName("COBRANCADEEVENTO").AsString
  vsNumeroBenefAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString

  If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then '1 -beneficiario 2 -outro
    If Not CurrentQuery.FieldByName("TITULARRESPONSAVEL").IsNull Then
      viHTitularResponsavelAnterior = CurrentQuery.FieldByName("TITULARRESPONSAVEL").AsInteger
    End If
  Else
    If Not CurrentQuery.FieldByName("PESSOARESPONSAVEL").IsNull Then
      viHTitularResponsavelAnterior = CurrentQuery.FieldByName("PESSOARESPONSAVEL").AsInteger
    End If
  End If

  Dim qSelect As Object
  Set qSelect = NewQuery

  qSelect.Clear
  qSelect.Add("SELECT LOCALFATURAMENTO")
  qSelect.Add("  FROM SAM_CONTRATO")
  qSelect.Add(" WHERE HANDLE = :CONTRATO")
  qSelect.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qSelect.Active = True

  If qSelect.FieldByName("LOCALFATURAMENTO").AsString = "F" Then
    'Se a família ainda não foi faturada, permitir alterar o dia de cobrança.
    qSelect.Clear
    qSelect.Add("SELECT COUNT(*) QTDREGISTROS")
    qSelect.Add("  FROM SFN_ROTINAFINFAT_FAMFAM")
    qSelect.Add("WHERE FAMILIA = :FAMILIA")
    qSelect.ParamByName("FAMILIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSelect.Active = True

    If qSelect.FieldByName("QTDREGISTROS").AsInteger = 0 Then
      DIACOBRANCA.ReadOnly       = False
      PROXIMOVENCIMENTO.ReadOnly = True
    Else
      DIACOBRANCA.ReadOnly       = True
      PROXIMOVENCIMENTO.ReadOnly = False
    End If
  End If

  Set qSelect = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  BOTAOCANCELAR.Visible              = False
  BOTAOFINANCEIRO.Visible            = False
  BOTAOREATIVAR.Visible              = False
  BOTAOCONSULTAVALORMODULO.Visible   = False
  BOTAOCRIAPESSOARESPONSAVEL.Visible = False
  BOTAODECLARACAO.Visible            = False
  BOTAOGRIDBENEFICIARIOS.Visible     = False
  BOTAOINSCRICAO.Visible             = False
  BOTAOMUDARRESPONSAVELFIN.Visible   = False
  vsModoEdicao = "I"

  FAMILIA.ReadOnly            = False
  PERIODOFATURAMENTO.ReadOnly = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface   As Object
  Dim viResultado As Integer
  Dim vsMensagem  As String

  Set Interface = CreateBennerObject("BSBEN020.Familia")

  'Os campos a seguir devem ter a propriedade ReadOnly modificada para "False"
  'para que seus valores possam ser alterados
  FAMILIA.ReadOnly             = False
  DATAULTIMOREAJUSTE.ReadOnly  = False
  DIACOBRANCAORIGINAL.ReadOnly = False
  PROXIMOVENCIMENTO.ReadOnly   = False
  DATABLOQUEIO.ReadOnly        = False
  CONVENIO.ReadOnly            = False

  viResultado = Interface.BeforePost(CurrentSystem, _
                                     CurrentQuery.TQuery, _
                                     viDiaCobrancaAnterior, _
                                     viTabResponsavelAnterior, _
                                     viHTitularResponsavelAnterior, _
                                     vsCobrancaDeEventoAnterior, _
                                     vsNumeroBenefAutomaticoAnterior, _
                                     vsMensagem)

  Set Interface = Nothing

  'Voltar os campos para ReadOnly
  FAMILIA.ReadOnly             = True
  DATAULTIMOREAJUSTE.ReadOnly  = True
  DIACOBRANCAORIGINAL.ReadOnly = True
  PROXIMOVENCIMENTO.ReadOnly   = True
  DATABLOQUEIO.ReadOnly        = True
  CONVENIO.ReadOnly            = True

  If viResultado = 1 Then
    CanContinue = False
    bsShowMessage(vsMensagem, "E")

    Dim qContrato As Object
    Set qContrato = NewQuery

    qContrato.Clear
    qContrato.Add("SELECT LOCALFATURAMENTO, NUMEROFAMILIAAUTOMATICO")
    qContrato.Add("  FROM SAM_CONTRATO")
    qContrato.Add(" WHERE HANDLE = :CONTRATO")
    qContrato.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
    qContrato.Active = True

    If qContrato.FieldByName("LOCALFATURAMENTO").AsString = "F" Then
      PERIODOFATURAMENTO.ReadOnly = False
      DIACOBRANCA.ReadOnly        = False
    End If

    'Se família automática protege o campo
    If (qContrato.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString <> "S") Then
      FAMILIA.ReadOnly= False
    End If

    Set qContrato = Nothing
  Else
    If vsMensagem <> "" Then
      bsShowMessage(vsMensagem, "I")
    End If
  End If

  'Se estiver em modo desktop a transação deve ser iniciada
  'Isto é necessário devido ao fato dos componentes BVirtual não controlarem transação
  If CurrentQuery.IsVirtual And _
     VisibleMode And (Not InTransaction) Then
    StartTransaction
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCRIAPESSOARESPONSAVEL"
			BOTAOCRIAPESSOARESPONSAVEL_OnClick
		Case "BOTAODECLARACAO"
			BOTAODECLARACAO_OnClick
		Case "BOTAOGRIDBENEFICIARIOS"
			BOTAOGRIDBENEFICIARIOS_OnClick
		Case "BOTAOINSCRICAO"
			BOTAOINSCRICAO_OnClick
		Case "BOTAOREATIVAR"
			BOTAOREATIVAR_OnClick
	End Select
End Sub

Public Sub TABLE_UpdateRequired()
 If CurrentQuery.FieldByName("CONVENIO").IsNull Then
   Dim qContrato As Object
   Set qContrato = NewQuery

   qContrato.Add("SELECT CONVENIO")
   qContrato.Add("FROM SAM_CONTRATO")
   qContrato.Add("WHERE HANDLE = :HCONTRATO")
   qContrato.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
   qContrato.Active = True
   CurrentQuery.FieldByName("CONVENIO").AsString = qContrato.FieldByName("CONVENIO").AsString

   Set qContrato = Nothing
 End If
End Sub

Public Sub TITULARRESPONSAVEL_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim qParamBenef     As Object
  Dim Interface       As Object
  Dim viHBeneficiario As Long

  Set qParamBenef = NewQuery

  qParamBenef.Add("SELECT CONSULTADETALHADACENTRAL")
  qParamBenef.Add("FROM SAM_PARAMETROSATENDIMENTO")
  qParamBenef.Active = True

  If qParamBenef.FieldByName("CONSULTADETALHADACENTRAL").AsString = "S" Then
    Set Interface = CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")

    viHBeneficiario = Interface.Filtro(CurrentSystem, _
                                       1, _
                                       "")
  Else
    Dim vsCampos As String
    Dim vsColunas As String
    Dim vsCriterio As String

    Set Interface = CreateBennerObject("Procura.Procurar")

    vsColunas  = "SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_BENEFICIARIO.EHTITULAR|SAM_MATRICULA.MATRICULA|SAM_MATRICULA.CPF|SAM_MATRICULA.RG"
    vsCampos   = "Nome|Beneficiário|Titular|Matrícula|CPF|RG"
    vsCriterio = "SAM_BENEFICIARIO.EHTITULAR = 'S'"

    viHBeneficiario = Interface.Exec(CurrentSystem, _
                                     "SAM_BENEFICIARIO|SAM_MATRICULA[SAM_BENEFICIARIO.MATRICULA = SAM_MATRICULA.HANDLE ]", _
                                     vsColunas, _
                                     1, _
                                     vsCampos, _
                                     vsCriterio, _
                                     "Procura por Beneficiarios", _
                                     False, _
                                     "")
  End If

  If viHBeneficiario <>0 Then
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("TITULARRESPONSAVEL").AsInteger = viHBeneficiario
  End If

  Set qParamBenef = Nothing
  Set Interface   = Nothing
End Sub

Public Function DiasPorMes(Ano As Integer, Mes As Integer)As Integer
  Dim vDiasMes(12)As Integer

  vDiasMes(1) = 31
  vDiasMes(2) = 28
  vDiasMes(3) = 31
  vDiasMes(4) = 30
  vDiasMes(5) = 31
  vDiasMes(6) = 30
  vDiasMes(7) = 31
  vDiasMes(8) = 31
  vDiasMes(9) = 30
  vDiasMes(10) = 31
  vDiasMes(11) = 30
  vDiasMes(12) = 31

  DiasPorMes = vDiasMes(Mes)

  If Mes = 2 Then
    If(Ano Mod 4 = 0)And((Ano Mod 100 <>0)Or(Ano Mod 400 = 0))Then
    DiasPorMes = 29
  End If
End If

End Function

Public Function RetornaContaFinanceira As Long
  Dim dllContaFin As Object
  Set dllContaFin = CreateBennerObject("Financeiro.Contafin")

  RetornaContaFinanceira = dllContaFin.Qual(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 8)

  Set dllContaFin = Nothing
End Function
