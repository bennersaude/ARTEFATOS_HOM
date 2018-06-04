'HASH: 24B681F513D585497211B9F28808B478
'SAM_SOLICITAUX_BENEFICIO
'#Uses "*bsShowMessage"

'Última alteração: Milton/17/01/2002 -SMS 5976

Option Explicit
Dim vGuarda As Integer
Dim interface As Object

'#Uses "*Arredonda"

Public Sub BOTAOADIANTAMENTO_OnClick()
  Dim SQL As Object
  Dim SQLDIA As Object
  Set SQLDIA = NewQuery
  Set SQL = NewQuery

  If Not CurrentQuery.FieldByName("VALORPRESTCONTAS").IsNull Then
    bsShowMessage("Não é possível efetuar o Adiantamento Provisírio" + Chr(13) + "Existe Prestação de Contas efetuada", "I")
    Exit Sub
  End If

  SQLDIA.Active = False
  SQLDIA.Clear
  SQLDIA.Add("SELECT SUM(QTDDIARIASSOLIC) AS DIARIAS FROM SAM_SOLICITAUX_BENEFICIO_DIA")
  SQLDIA.Add("WHERE SOLICITAUXBENEFICIO = :HSOLICITAUXBEN")
  SQLDIA.ParamByName("HSOLICITAUXBEN").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLDIA.Active = True

  If SQLDIA.FieldByName("DIARIAS").IsNull Then
    bsShowMessage("Os valores das diárias estão em branco", "I")
    SQLDIA.Active = False
    Set SQLDIA = Nothing
    Exit Sub
  Else
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT TABCLASSIFICACAO, SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
    SQL.ParamByName("HSOLICITAUX").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
    SQL.Active = True

    If SQL.FieldByName("SITUACAO").AsString = "L" Then
      If(SQL.FieldByName("TABCLASSIFICACAO").AsInteger = 2)Then
      Set interface = CreateBennerObject("SamSolicitAux.AdiantamentoProvisorio")
      Set interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      RefreshNodesWithTable("SAM_SOLICITAUX_BENEFICIO")
      Set interface = Nothing
      SQL.Active = False
      SQLDIA.Active = False
      Set SQLDIA = Nothing
      Set SQL = Nothing
    Else
      bsShowMessage("Este tipo de benefício não permite adiantamento provisório", "I")
      SQL.Active = False
      Set SQL = Nothing
      Exit Sub
    End If
  Else
    bsShowMessage("Somente é possível efetuar o Adiantamento Provisório se a solicitação estiver LIBERADA", "I")
    SQL.Active = False
    Set SQL = Nothing
  End If
End If
End Sub

Public Sub BOTAOPRESTCONTAS_OnClick()
  Dim SQLDIA As Object
  Set SQLDIA = NewQuery

  If(Not CurrentQuery.FieldByName("VALORPRESTCONTAS").IsNull)Then
  bsShowMessage("Prestação de Contas já efetuada", "I")
  Exit Sub
End If

SQLDIA.Clear
SQLDIA.Add("SELECT SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
SQLDIA.ParamByName("HSOLICITAUX").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
SQLDIA.Active = True

If SQLDIA.FieldByName("SITUACAO").AsString <>"L" Then
  bsShowMessage("Somente é possível efetuar a Prestação de Contas se a solicitação estiver LIBERADA", "I")
  SQLDIA.Active = False
  Set SQLDIA = Nothing
  Exit Sub
End If

If CurrentQuery.FieldByName("VALORADIANTAMENTOPROV").IsNull Then
  bsShowMessage("Adiantamento Provisório não efetuado", "I")
  Exit Sub
Else
  SQLDIA.Active = False
  SQLDIA.Clear
  SQLDIA.Add(" SELECT QTDDIARIASPRESTCONTAS,VALORHOSPEDAGEMPRESTCONTAS,VALORREFEICAOPRESTCONTAS")
  SQLDIA.Add(" FROM SAM_SOLICITAUX_BENEFICIO_DIA ")
  SQLDIA.Add(" WHERE SOLICITAUXBENEFICIO = :SOLICITAUXBEN ")
  SQLDIA.Add(" AND VALORHOSPEDAGEMPRESTCONTAS <> 0 ")
  SQLDIA.ParamByName("SOLICITAUXBEN").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLDIA.Active = True
  If SQLDIA.EOF Then
    bsShowMessage("Não foi possível realizar esta operação,os valores" + Chr(13) + "da prestação de contas estão em aberto", "I")
    Exit Sub
  Else
    Set interface = CreateBennerObject("SamSolicitAux.PrestacaoContas")
    Set interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    RefreshNodesWithTable("SAM_SOLICITAUX_BENEFICIO")
    Set interface = Nothing
  End If
End If

SQLDIA.Active = False
Set SQLDIA = Nothing

End Sub


Public Sub BOTAORECALCULAR_OnClick()
  If CurrentQuery.State = 1 Then
    If Not(CurrentQuery.FieldByName("EVENTO").IsNull)And _
           Not(CurrentQuery.FieldByName("GRAU").IsNull)Then
      Dim Peg As Object
      Dim vPrecoBeneficio As Double
      Dim SQL As Object
      Dim vDataSolicitacao As Date
      Dim vBeneficiario As Long
      Dim vEstado As Long
      Dim vFilial As Long
      Dim vMunicipio As Long
      Dim vValorPrimeira As Double
      Dim vValorSegunda As Double
      Dim vValorDemais As Double
      Dim vXTHM As Long
      Dim vCodigoPagto As Long
      Dim vPacoteLimiteTipo As String
      Dim vPacoteLimiteValor As Double



      Set SQL = NewQuery

      SQL.Clear
      SQL.Add("SELECT A.BENEFICIARIO, A.ESTADO, A.MUNICIPIO, A.CODIGO, A.PACOTEAUXILIO,")
      SQL.Add("       B.LIMITETIPO, B.LIMITEVALOR, A.FILIAL ")
      SQL.Add("FROM SAM_SOLICITAUX A, SAM_PACOTEAUXILIO B")
      SQL.Add("WHERE A.HANDLE = :SOLICITAUXCODIGO")
      SQL.Add("  AND B.HANDLE = A.PACOTEAUXILIO")
      SQL.ParamByName("SOLICITAUXCODIGO").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
      SQL.Active = True

      vDataSolicitacao = CurrentQuery.FieldByName("DATACADASTRO").AsDateTime
      vBeneficiario = SQL.FieldByName("BENEFICIARIO").AsInteger
      vEstado = SQL.FieldByName("ESTADO").AsInteger
      vFilial = SQL.FieldByName("FILIAL").AsInteger
      vMunicipio = SQL.FieldByName("MUNICIPIO").AsInteger
      vPacoteLimiteTipo = SQL.FieldByName("LIMITETIPO").AsString
      vPacoteLimiteValor = SQL.FieldByName("LIMITEVALOR").AsFloat

      SQL.Clear
      SQL.Add("SELECT CODIGOXTHM, CODIGOPAGTO")
      SQL.Add("FROM SAM_PARAMETROSPROCCONTAS")
      SQL.Active = True

      vXTHM = SQL.FieldByName("CODIGOXTHM").AsInteger
      vCodigoPagto = SQL.FieldByName("CODIGOPAGTO").AsInteger

      Set Peg = CreateBennerObject("SamPeg.Rotinas")

      Peg.Inicializar(CurrentSystem)

      vPrecoBeneficio = Peg.PegaPreco(CurrentSystem, _
                        CurrentQuery.FieldByName("EVENTO").AsInteger, _
                        CurrentQuery.FieldByName("GRAU").AsInteger, _
                        vBeneficiario, _
                        0, _
                        0, _
                        vFilial, _
                        vMunicipio, _
                        vEstado, _
                        Null, _
                        vDataSolicitacao, _
                        vCodigoPagto, _
                        CurrentQuery.FieldByName("QUANTIDADE").AsFloat, _
                        vXTHM, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
						Now, _
						False, _
						"1", _
                        vValorPrimeira, _
                        vValorSegunda, _
                        vValorDemais)

      SQL.Clear
      SQL.Add("UPDATE SAM_SOLICITAUX_BENEFICIO SET")
      SQL.Add("   VALOREVENTO = :VALOREVENTO,")
      SQL.Add("   VALORUNITARIOLIBERADO = :VALORUNITARIOLIBERADO,")

      If vPacoteLimiteTipo = "F" Then
        SQL.Add("   VALORLIMITEEVENTO = :VALORLIMITEEVENTO")
        SQL.ParamByName("VALORLIMITEEVENTO").Value = vValorPrimeira * vPacoteLimiteValor
      Else
        SQL.Add("   VALORLIMITEEVENTO = NULL")
      End If

      If Not InTransaction Then StartTransaction

      SQL.Add("WHERE HANDLE = :HSOLICITAUXBENEFICIO")
      SQL.ParamByName("HSOLICITAUXBENEFICIO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.ParamByName("VALOREVENTO").Value = vValorPrimeira
      SQL.ParamByName("VALORUNITARIOLIBERADO").Value = vValorPrimeira
      SQL.ExecSQL

      If InTransaction Then Commit

      Peg.Finalizar

      Set Peg = Nothing
      Set SQL = Nothing

      RefreshNodesWithTable("SAM_SOLICITAUX_BENEFICIO")
    Else
      bsShowMessage("O registro não pode estar em edição", "I")
    End If
  End If
End Sub

Public Sub EVENTO_OnExit()

  If CurrentQuery.State <>1 Then
    If Not(CurrentQuery.FieldByName("EVENTO").IsNull)And _
           Not(CurrentQuery.FieldByName("GRAU").IsNull)Then
      Dim Peg As Object
      Dim vPrecoBeneficio As Double
      Dim SQL As Object
      Dim vDataSolicitacao As Date
      Dim vBeneficiario As Long
      Dim vEstado As Long
      Dim vFilial As Long
      Dim vMunicipio As Long
      Dim vValorPrimeira As Double
      Dim vValorDemais As Double
      Dim vXTHM As Long
      Dim vCodigoPagto As Long
      Dim vPacoteLimiteTipo As String
      Dim vPacoteLimiteValor As Double



      Set SQL = NewQuery

      SQL.Clear
      SQL.Add("SELECT A.BENEFICIARIO, A.ESTADO, A.MUNICIPIO, A.CODIGO, A.PACOTEAUXILIO,")
      SQL.Add("       B.LIMITETIPO, B.LIMITEVALOR, A.FILIAL ")
      SQL.Add("FROM SAM_SOLICITAUX A, SAM_PACOTEAUXILIO B")
      SQL.Add("WHERE A.HANDLE = :SOLICITAUXCODIGO")
      SQL.Add("  AND B.HANDLE = A.PACOTEAUXILIO")
      SQL.ParamByName("SOLICITAUXCODIGO").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
      SQL.Active = True

      vDataSolicitacao = CurrentQuery.FieldByName("DATACADASTRO").AsDateTime
      vBeneficiario = SQL.FieldByName("BENEFICIARIO").AsInteger
      vEstado = SQL.FieldByName("ESTADO").AsInteger
      vFilial = SQL.FieldByName("FILIAL").AsInteger
      vMunicipio = SQL.FieldByName("MUNICIPIO").AsInteger
      vPacoteLimiteTipo = SQL.FieldByName("LIMITETIPO").AsString
      vPacoteLimiteValor = SQL.FieldByName("LIMITEVALOR").AsFloat

      SQL.Clear
      SQL.Add("SELECT CODIGOXTHM, CODIGOPAGTO")
      SQL.Add("FROM SAM_PARAMETROSPROCCONTAS")
      SQL.Active = True

      vXTHM = SQL.FieldByName("CODIGOXTHM").AsInteger
      vCodigoPagto = SQL.FieldByName("CODIGOPAGTO").AsInteger

      Set Peg = CreateBennerObject("SamPeg.Rotinas")

      Peg.Inicializar(CurrentSystem)

      vPrecoBeneficio = Peg.PegaPreco(CurrentSystem, _
                        CurrentQuery.FieldByName("EVENTO").AsInteger, _
                        CurrentQuery.FieldByName("GRAU").AsInteger, _
                        vBeneficiario, _
                        0, _
                        0, _
                        vFilial, _
                        vMunicipio, _
                        vEstado, _
                        Null, _
                        vDataSolicitacao, _
                        vCodigoPagto, _
                        CurrentQuery.FieldByName("QUANTIDADE").AsFloat, _
                        vXTHM, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
						Now, _
						False, _
						"1", _
                        vValorPrimeira, _
                        0, _
                        vValorDemais)

      CurrentQuery.FieldByName("VALOREVENTO").Value = vValorPrimeira
      CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat = vValorPrimeira

      If vPacoteLimiteTipo = "F" Then
        CurrentQuery.FieldByName("VALORLIMITEEVENTO").AsFloat = vValorPrimeira * vPacoteLimiteValor
      Else
        CurrentQuery.FieldByName("VALORLIMITEEVENTO").Clear
      End If

      Peg.Finalizar

      Set Peg = Nothing
      Set SQL = Nothing
    Else
      CurrentQuery.FieldByName("VALOREVENTO").Clear
      CurrentQuery.FieldByName("VALORLIMITEEVENTO").Clear
      CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").Clear
    End If
  End If
End Sub

Public Sub GRAU_OnExit()

  If CurrentQuery.State <>1 Then
    If Not(CurrentQuery.FieldByName("EVENTO").IsNull)And _
           Not(CurrentQuery.FieldByName("GRAU").IsNull)Then
      Dim Peg As Object
      Dim vPrecoBeneficio As Double
      Dim SQL As Object
      Dim vDataSolicitacao As Date
      Dim vBeneficiario As Long
      Dim vEstado As Long
      Dim vFilial As Long
      Dim vMunicipio As Long
      Dim vValorPrimeira As Double
      Dim vValorDemais As Double
      Dim vXTHM As Long
      Dim vCodigoPagto As Long
      Dim vPacoteLimiteTipo As String
      Dim vPacoteLimiteValor As Double



      Set SQL = NewQuery

      SQL.Clear
      SQL.Add("SELECT A.BENEFICIARIO, A.ESTADO, A.MUNICIPIO, A.CODIGO, A.PACOTEAUXILIO,")
      SQL.Add("       B.LIMITETIPO, B.LIMITEVALOR, A.FILIAL ")
      SQL.Add("FROM SAM_SOLICITAUX A, SAM_PACOTEAUXILIO B")
      SQL.Add("WHERE A.HANDLE = :SOLICITAUXCODIGO")
      SQL.Add("  AND B.HANDLE = A.PACOTEAUXILIO")
      SQL.ParamByName("SOLICITAUXCODIGO").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
      SQL.Active = True

      vDataSolicitacao = CurrentQuery.FieldByName("DATACADASTRO").AsDateTime
      vBeneficiario = SQL.FieldByName("BENEFICIARIO").AsInteger
      vEstado = SQL.FieldByName("ESTADO").AsInteger
      vFilial = SQL.FieldByName("FILIAL").AsInteger
      vMunicipio = SQL.FieldByName("MUNICIPIO").AsInteger
      vPacoteLimiteTipo = SQL.FieldByName("LIMITETIPO").AsString
      vPacoteLimiteValor = SQL.FieldByName("LIMITEVALOR").AsFloat

      SQL.Clear
      SQL.Add("SELECT CODIGOXTHM, CODIGOPAGTO")
      SQL.Add("FROM SAM_PARAMETROSPROCCONTAS")
      SQL.Active = True

      vXTHM = SQL.FieldByName("CODIGOXTHM").AsInteger
      vCodigoPagto = SQL.FieldByName("CODIGOPAGTO").AsInteger

      Set Peg = CreateBennerObject("SamPeg.Rotinas")

      Peg.Inicializar(CurrentSystem)

      vPrecoBeneficio = Peg.PegaPreco(CurrentSystem, _
                        CurrentQuery.FieldByName("EVENTO").AsInteger, _
                        CurrentQuery.FieldByName("GRAU").AsInteger, _
                        vBeneficiario, _
                        0, _
                        0, _
                        vFilial, _
                        vMunicipio, _
                        vEstado, _
                        Null, _
                        vDataSolicitacao, _
                        vCodigoPagto, _
                        CurrentQuery.FieldByName("QUANTIDADE").AsFloat, _
                        vXTHM, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
						Now, _
						False, _
						"1", _
                        vValorPrimeira, _
                        0, _
                        vValorDemais)

      CurrentQuery.FieldByName("VALOREVENTO").Value = vValorPrimeira
      CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat = vValorPrimeira

      If vPacoteLimiteTipo = "F" Then
        CurrentQuery.FieldByName("VALORLIMITEEVENTO").AsFloat = vValorPrimeira * vPacoteLimiteValor
        'André - SMS 28326 - 15/10/2004
        CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat = vValorPrimeira * vPacoteLimiteValor
        'FIM SMS 28326
      Else
        CurrentQuery.FieldByName("VALORLIMITEEVENTO").Clear
      End If

      Peg.Finalizar

      Set Peg = Nothing
      Set SQL = Nothing
    Else
      CurrentQuery.FieldByName("VALOREVENTO").Clear
      CurrentQuery.FieldByName("VALORLIMITEEVENTO").Clear
      CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").Clear
    End If
  End If
End Sub

Public Sub QUANTIDADE_OnExit()

  If CurrentQuery.State <>1 Then
    If Not(CurrentQuery.FieldByName("EVENTO").IsNull)And _
           Not(CurrentQuery.FieldByName("GRAU").IsNull)Then
      Dim Peg As Object
      Dim vPrecoBeneficio As Double
      Dim SQL As Object
      Dim vDataSolicitacao As Date
      Dim vBeneficiario As Long
      Dim vEstado As Long
      Dim vFilial As Long
      Dim vMunicipio As Long
      Dim vValorPrimeira As Double
      Dim vValorDemais As Double
      Dim vXTHM As Long
      Dim vCodigoPagto As Long
      Dim vPacoteLimiteTipo As String
      Dim vPacoteLimiteValor As Double



      Set SQL = NewQuery

      SQL.Clear
      SQL.Add("SELECT A.BENEFICIARIO, A.ESTADO, A.MUNICIPIO, A.CODIGO, A.PACOTEAUXILIO,")
      SQL.Add("       B.LIMITETIPO, B.LIMITEVALOR, A.FILIAL ")
      SQL.Add("FROM SAM_SOLICITAUX A, SAM_PACOTEAUXILIO B")
      SQL.Add("WHERE A.HANDLE = :SOLICITAUXCODIGO")
      SQL.Add("  AND B.HANDLE = A.PACOTEAUXILIO")
      SQL.ParamByName("SOLICITAUXCODIGO").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
      SQL.Active = True

      vDataSolicitacao = CurrentQuery.FieldByName("DATACADASTRO").AsDateTime
      vBeneficiario = SQL.FieldByName("BENEFICIARIO").AsInteger
      vEstado = SQL.FieldByName("ESTADO").AsInteger
      vFilial = SQL.FieldByName("FILIAL").AsInteger
      vMunicipio = SQL.FieldByName("MUNICIPIO").AsInteger
      vPacoteLimiteTipo = SQL.FieldByName("LIMITETIPO").AsString
      vPacoteLimiteValor = SQL.FieldByName("LIMITEVALOR").AsFloat

      SQL.Clear
      SQL.Add("SELECT CODIGOXTHM, CODIGOPAGTO")
      SQL.Add("FROM SAM_PARAMETROSPROCCONTAS")
      SQL.Active = True

      vXTHM = SQL.FieldByName("CODIGOXTHM").AsInteger
      vCodigoPagto = SQL.FieldByName("CODIGOPAGTO").AsInteger

      Set Peg = CreateBennerObject("SamPeg.Rotinas")

      Peg.Inicializar(CurrentSystem)

      vPrecoBeneficio = Peg.PegaPreco(CurrentSystem, _
                        CurrentQuery.FieldByName("EVENTO").AsInteger, _
                        CurrentQuery.FieldByName("GRAU").AsInteger, _
                        vBeneficiario, _
                        0, _
                        0, _
                        vFilial, _
                        vMunicipio, _
                        vEstado, _
                        Null, _
                        vDataSolicitacao, _
                        vCodigoPagto, _
                        CurrentQuery.FieldByName("QUANTIDADE").AsFloat, _
                        vXTHM, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
                        0, _
						Now, _
						False, _
						"1", _
                        vValorPrimeira, _
                        0, _
                        vValorDemais)

      CurrentQuery.FieldByName("VALOREVENTO").Value = vValorPrimeira
      CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat = vValorPrimeira

      If vPacoteLimiteTipo = "F" Then
        CurrentQuery.FieldByName("VALORLIMITEEVENTO").AsFloat = vValorPrimeira * vPacoteLimiteValor
        'André - SMS 28326 - 15/10/2004
        CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat = vValorPrimeira * vPacoteLimiteValor
        'FIM SMS 28326
      Else
        CurrentQuery.FieldByName("VALORLIMITEEVENTO").Clear
      End If

      Peg.Finalizar

      Set Peg = Nothing
      Set SQL = Nothing
    Else
      CurrentQuery.FieldByName("VALOREVENTO").Clear
      CurrentQuery.FieldByName("VALORLIMITEEVENTO").Clear
      CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").Clear
    End If
  End If
End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
  SQL.ParamByName("HSOLICITAUX").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString <>"A" Then
    EVENTO.ReadOnly = True
    GRAU.ReadOnly = True
    ESTADOORIGEM.ReadOnly = True
    ESTADODESTINO.ReadOnly = True
    MUNICIPIOORIGEM.ReadOnly = True
    MUNICIPIODESTINO.ReadOnly = True
    DATACADASTRO.ReadOnly = True
    QTDACOMPANHANTES.ReadOnly = True
    QUANTIDADE.ReadOnly = True
    'VALORREQUERIDO.ReadOnly =True
    'VALORNAOABONAVEL.ReadOnly =True
  Else
    EVENTO.ReadOnly = False
    GRAU.ReadOnly = False
    ESTADOORIGEM.ReadOnly = False
    ESTADODESTINO.ReadOnly = False
    MUNICIPIOORIGEM.ReadOnly = False
    MUNICIPIODESTINO.ReadOnly = False
    DATACADASTRO.ReadOnly = False
    QTDACOMPANHANTES.ReadOnly = False
    QUANTIDADE.ReadOnly = False
    'VALORREQUERIDO.ReadOnly =False
    'VALORNAOABONAVEL.ReadOnly =False
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vGuarda = 0
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  vGuarda = 1
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLPAC As Object
  Dim SQLSOL As Object
  Dim SQLEVE As Object
  Dim SQLBEN As Object
  Dim vDataSolicitacao As Date
  Dim vBeneficiario As Long
  Dim vEstado As Long
  Dim vMunicipio As Long
  Dim vTabClassificacao As Long

  Set SQLPAC = NewQuery
  Set SQLSOL = NewQuery
  Set SQLEVE = NewQuery
  Set SQLBEN = NewQuery

  SQLSOL.Active = False
  SQLSOL.Clear
  SQLSOL.Add(" SELECT BENEFICIARIO, ESTADO, MUNICIPIO, CODIGO, PACOTEAUXILIO, TABCLASSIFICACAO FROM SAM_SOLICITAUX ")
  SQLSOL.Add(" WHERE HANDLE = :SOLICITAUXCODIGO ")
  SQLSOL.ParamByName("SOLICITAUXCODIGO").AsInteger = CurrentQuery.FieldByName("SOLICITAUX").AsInteger
  SQLSOL.Active = True

  vDataSolicitacao = CurrentQuery.FieldByName("DATACADASTRO").AsDateTime
  vBeneficiario = SQLSOL.FieldByName("BENEFICIARIO").AsInteger
  vEstado = SQLSOL.FieldByName("ESTADO").AsInteger
  vMunicipio = SQLSOL.FieldByName("MUNICIPIO").AsInteger
  vTabClassificacao = SQLSOL.FieldByName("TABCLASSIFICACAO").AsInteger

  If vTabClassificacao = 1 Then
    If CurrentQuery.FieldByName("EVENTO").IsNull Then
      CanContinue = False
      bsShowMessage("O evento deve ser informado", "E")
      Exit Sub
    End If

    If CurrentQuery.FieldByName("GRAU").IsNull Then
      CanContinue = False
      bsShowMessage("O grau deve ser informado", "E")
      Exit Sub
    End If

    If CurrentQuery.FieldByName("QUANTIDADE").AsFloat = 0 Then
      CanContinue = False
      bsShowMessage("Deve ser informado um valor maior que ZERO para a quantidade", "E")
      Exit Sub
    End If

    SQLPAC.Active = False
    SQLPAC.Clear
    SQLPAC.Add(" SELECT LIMITETIPO, LIMITEVALOR,HANDLE FROM SAM_PACOTEAUXILIO ")
    SQLPAC.Add(" WHERE HANDLE = :PACOTEAUXILIO ")
    SQLPAC.ParamByName("PACOTEAUXILIO").AsInteger = SQLSOL.FieldByName("PACOTEAUXILIO").AsInteger
    SQLPAC.Active = True

    'SMS 73656 - Marcelo Barbosa - 14/12/2006
    If CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat = 0 And _
       (SQLPAC.FieldByName("LIMITETIPO").AsString = "F" Or _
       SQLPAC.FieldByName("LIMITETIPO").AsString = "V") Then
      CanContinue = False
      bsShowMessage("Deve ser informado um valor maior que ZERO para o valor unitário liberado", "E")
      Exit Sub
    End If

    SQLEVE.Active = False
    SQLEVE.Clear
    SQLEVE.Add(" SELECT EVENTO,GRAU FROM SAM_PACOTEAUXILIO_EVENTOS ")
    SQLEVE.Add(" WHERE PACOTEAUXLIO = :PACOTEAUX ")
    SQLEVE.ParamByName("PACOTEAUX").AsInteger = SQLSOL.FieldByName("PACOTEAUXILIO").AsInteger
    SQLEVE.Active = True


    While(Not SQLEVE.EOF)And(CurrentQuery.FieldByName("EVENTO").AsInteger <>SQLEVE.FieldByName("EVENTO").AsInteger)
    SQLEVE.Next
  Wend
  If(CurrentQuery.FieldByName("EVENTO").AsInteger <>SQLEVE.FieldByName("EVENTO").AsInteger)Then
  bsShowMessage("Evento não cadastrado no Pacote", "I")
  'EVENTO.SetFocus
  CanContinue = False
  Exit Sub
End If

While(Not SQLEVE.EOF)And(CurrentQuery.FieldByName("GRAU").AsInteger <>SQLEVE.FieldByName("GRAU").AsInteger)
SQLEVE.Next
Wend
If(CurrentQuery.FieldByName("GRAU").AsInteger <>SQLEVE.FieldByName("GRAU").AsInteger)Then
bsShowMessage("Grau não cadastrado no Pacote", "E")
'GRAU.SetFocus
CanContinue = False
Exit Sub
End If
SQLEVE.Next
Else
  If CurrentQuery.FieldByName("ESTADOORIGEM").IsNull Then
    CanContinue = False
    bsShowMessage("O estado origem deve ser informado", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("MUNICIPIOORIGEM").IsNull Then
    CanContinue = False
    bsShowMessage("O município origem deve ser informado", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("ESTADODESTINO").IsNull Then
    CanContinue = False
    bsShowMessage("O estado destino deve ser informado", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("MUNICIPIODESTINO").IsNull Then
    CanContinue = False
    bsShowMessage("O município destino deve ser informado", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("QTDACOMPANHANTES").IsNull Then
    CanContinue = False
    bsShowMessage("A quantidade de acompanhantes deve ser informada", "E")
    Exit Sub
  End If
End If

If(vTabClassificacao = 2)Then
If CurrentQuery.FieldByName("MUNICIPIOORIGEM").Value = CurrentQuery.FieldByName("MUNICIPIODESTINO").Value Then
  bsShowMessage("O Município de Origem não pode ser igual ao Município de Destino", "E")
  CanContinue = False
  Exit Sub
End If

Else
  If Not(CurrentQuery.FieldByName("VALORLIMITEEVENTO").IsNull)And _
         (CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat >CurrentQuery.FieldByName("VALORLIMITEEVENTO").AsFloat)Then
    CanContinue = False
    bsShowMessage("Valor unitário liberado não pode ser superior ao valor limite do evento", "E")
    Exit Sub
  End If

  If(SQLPAC.FieldByName("LIMITETIPO").AsString = "Q")And _
     (CurrentQuery.FieldByName("QUANTIDADE").AsFloat >SQLPAC.FieldByName("LIMITEVALOR").AsFloat)Then
  CanContinue = False
  bsShowMessage("Quantidade superior à quantidade permitida na configuração do pacote", "E")
  Exit Sub
End If

'If(CurrentQuery.FieldByName("TABCLASSIFICACAO").AsInteger =2)And _
'  (Arredonda(CurrentQuery.FieldByName("QUANTIDADE").AsFloat * CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat)>CurrentQuery.FieldByName("VALORREQUERIDO").AsFloat)Then
'   CanContinue=False
'   MsgBox("O cálculo 'Valor unitário liberado * Quantidade' não pode ultrapassar o valor requerido")
'   Exit Sub
'End If
End If

SQLSOL.Active = False
SQLPAC.Active = False
Set SQLSOL = Nothing
Set SQLPAC = Nothing

'Dim vValorTotalLiberado As Double
'If CurrentQuery.FieldByName("TABCLASSIFICACAO").AsInteger =2 Then
'   vValorTotalLiberado=Arredonda(CurrentQuery.FieldByName("QUANTIDADE").AsFloat * CurrentQuery.FieldByName("VALORUNITARIOLIBERADO").AsFloat)
'   CurrentQuery.FieldByName("VALORNAOABONAVEL").AsFloat=CurrentQuery.FieldByName("VALORREQUERIDO").AsFloat -vValorTotalLiberado
'End If

End Sub


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEventoPacote(True)' só último nível
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraGrauPacote
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
End Sub

Public Function ProcuraEventoPacote(pUltimoNivel As Boolean)As Long
  'Balani SMS 55210 23/12/2005
  Dim interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Active =False
  SQL.Clear
  SQL.Add(" SELECT PACOTEAUXILIO FROM SAM_SOLICITAUX ")
  SQL.Add(" WHERE HANDLE = :SOLICITAUXCODIGO ")
  SQL.ParamByName("SOLICITAUXCODIGO").AsInteger =CurrentQuery.FieldByName("SOLICITAUX").AsInteger
  SQL.Active =True

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|Z_DESCRICAO|NIVELAUTORIZACAO"

  If pUltimoNivel Then
    vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
  Else
    vCriterio = "SAM_TGE.HANDLE > 0"
  End If

  vCampos = "Evento|Descrição|Nível"

  If CurrentQuery.FieldByName("GRAU").IsNull Then
    ProcuraEventoPacote =interface.Exec(CurrentSystem,"SAM_TGE|SAM_PACOTEAUXILIO_EVENTOS[SAM_PACOTEAUXILIO_EVENTOS.EVENTO = SAM_TGE.HANDLE AND SAM_PACOTEAUXILIO_EVENTOS.PACOTEAUXLIO = " + SQL.FieldByName("PACOTEAUXILIO").AsString + "]",vColunas,2,vCampos,vCriterio, "Tabela Geral de Eventos",True,EVENTO.Text)
  Else
    ProcuraEventoPacote =interface.Exec(CurrentSystem,"SAM_TGE|SAM_PACOTEAUXILIO_EVENTOS[SAM_PACOTEAUXILIO_EVENTOS.EVENTO = SAM_TGE.HANDLE AND SAM_PACOTEAUXILIO_EVENTOS.PACOTEAUXLIO = " + SQL.FieldByName("PACOTEAUXILIO").AsString + " AND SAM_PACOTEAUXILIO_EVENTOS.GRAU = " + CurrentQuery.FieldByName("GRAU").AsString + "]",vColunas,2,vCampos,vCriterio, "Tabela Geral de Eventos",True,EVENTO.Text)
  End If

  Set SQL = Nothing
  Set interface = Nothing
  'Final SMS 55210
End Function

Public Function ProcuraGrauPacote()As Long
  'Balani SMS 55210 23/12/2005
  Dim interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Active =False
  SQL.Clear
  SQL.Add(" SELECT PACOTEAUXILIO FROM SAM_SOLICITAUX ")
  SQL.Add(" WHERE HANDLE = :SOLICITAUXCODIGO ")
  SQL.ParamByName("SOLICITAUXCODIGO").AsInteger =CurrentQuery.FieldByName("SOLICITAUX").AsInteger
  SQL.Active =True

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO"

  vCriterio = "SAM_GRAU.HANDLE = SAM_PACOTEAUXILIO_EVENTOS.GRAU"
  vCampos = "Código do Grau|Descrição|Tipo do Grau"


  If CurrentQuery.FieldByName("EVENTO").IsNull Then
    ProcuraGrauPacote =interface.Exec(CurrentSystem,"SAM_GRAU|SAM_PACOTEAUXILIO_EVENTOS[SAM_PACOTEAUXILIO_EVENTOS.GRAU = SAM_GRAU.HANDLE AND SAM_PACOTEAUXILIO_EVENTOS.PACOTEAUXLIO = " + SQL.FieldByName("PACOTEAUXILIO").AsString + "]",vColunas,2,vCampos,vCriterio,"Graus de Atuação",True,"")
  Else
    ProcuraGrauPacote =interface.Exec(CurrentSystem,"SAM_GRAU|SAM_PACOTEAUXILIO_EVENTOS[SAM_PACOTEAUXILIO_EVENTOS.GRAU = SAM_GRAU.HANDLE AND SAM_PACOTEAUXILIO_EVENTOS.PACOTEAUXLIO = " + SQL.FieldByName("PACOTEAUXILIO").AsString + " AND SAM_PACOTEAUXILIO_EVENTOS.EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + "]",vColunas,2,vCampos,vCriterio,"Graus de Atuação",True,"")
  End If

  Set SQL = Nothing
  Set interface = Nothing

  'final SMS 55210
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOADIANTAMENTO"
			BOTAOADIANTAMENTO_OnClick
		Case "BOTAOPRESTCONTAS"
			BOTAOPRESTCONTAS_OnClick
		Case "BOTAORECALCULAR"
			BOTAORECALCULAR_OnClick
	End Select
End Sub
