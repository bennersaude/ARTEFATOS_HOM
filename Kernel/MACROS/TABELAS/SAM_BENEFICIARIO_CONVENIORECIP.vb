'HASH: 222241B246E0B89AF606A5B70EE8289C
'#Uses "*bsShowMessage"
'Macro SAM_BENEFICIARIO_CONVENIORECIP

Dim vgSituacao As String

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim INTERFACE As Object
  Dim vHandle As Long
  Dim vCabecs As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String
  Dim vTitulo As String

  If (PRESTADOR.PopupCase <> 0) Then
    ShowPopup = False
    Set INTERFACE = CreateBennerObject("Procura.Procurar")

    vCabecs = "Código|Prestador|CPFCNPJ"
    vColunas = "PRESTADOR|NOME|CPFCNPJ"
    vCriterio = "SAM_PRESTADOR.CONVENIORECIPROCIDADE = 'S' "
    vTabela = "SAM_PRESTADOR"
    vTitulo = "Prestadores - Convênio de Reciprocidade"

    vHandle = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, False, "")

    If vHandle <> 0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("PRESTADOR").AsInteger = vHandle
    End If
    Set INTERFACE = Nothing
  Else
    ShowPopup = True
  End If
End Sub

Public Sub TABLE_AfterInsert()

  Dim QryMatricula As Object
  Dim SQL As Object
  Set QryMatricula = NewQuery
  Set SQL = NewQuery

  SQL.Add("SELECT MATRICULA ")
  SQL.Add("  FROM SAM_BENEFICIARIO ")
  SQL.Add(" WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True

  QryMatricula.Add("SELECT NOME, ")
  QryMatricula.Add("       CPF, ")
  QryMatricula.Add("       DATANASCIMENTO, ")
  QryMatricula.Add("       SEXO, ")
  QryMatricula.Add("       NOMEPAI, ")
  QryMatricula.Add("       NOMEMAE, ")
  QryMatricula.Add("       CARTAONACIONALSAUDE ")
  QryMatricula.Add("  FROM SAM_MATRICULA ")
  QryMatricula.Add(" WHERE HANDLE = :HANDLE ")
  QryMatricula.ParamByName("HANDLE").AsInteger = SQL.FieldByName("MATRICULA").AsInteger
  QryMatricula.Active = True

  CurrentQuery.FieldByName("NOME").AsString = QryMatricula.FieldByName("NOME").AsString
  CurrentQuery.FieldByName("CPF").AsString = QryMatricula.FieldByName("CPF").AsString
  CurrentQuery.FieldByName("DATANASCIMENTO").AsString = QryMatricula.FieldByName("DATANASCIMENTO").AsString
  CurrentQuery.FieldByName("SEXO").AsString = QryMatricula.FieldByName("SEXO").AsString
  CurrentQuery.FieldByName("NOMEPAI").AsString = QryMatricula.FieldByName("NOMEPAI").AsString
  CurrentQuery.FieldByName("NOMEMAE").AsString = QryMatricula.FieldByName("NOMEMAE").AsString
  CurrentQuery.FieldByName("CARTAONACIONALSAUDE").AsString = QryMatricula.FieldByName("CARTAONACIONALSAUDE").AsString
  CurrentQuery.FieldByName("SITUACAO").AsString = "P"

  Set QryMatricula = Nothing
  Set SQL = Nothing

End Sub

Public Sub TABLE_AfterScroll()

   If (VisibleMode And NodeInternalCode = 99) Or (WebMode And (WebMenuCode = "T1103" Or WebMenuCode = "T2886" Or WebMenuCode = "T1613" Or WebMenuCode = ""))Then
    If (Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull) Then
      CODIGO.ReadOnly = True
      MOTIVO.ReadOnly = True
      DATAINICIAL.ReadOnly = True
      DATAFINAL.ReadOnly = False
      'DATACANCELAMENTO.ReadOnly = False
      PRESTADOR.ReadOnly = True
      CODIGO.ReadOnly = True
    Else
      CODIGO.ReadOnly = False
      MOTIVO.ReadOnly = False
      DATAINICIAL.ReadOnly = False
      DATAFINAL.ReadOnly = False
      PRESTADOR.ReadOnly = False
      CODIGO.ReadOnly = False
    End If

    If (CurrentQuery.FieldByName("SITUACAO").AsString = "P" And CurrentQuery.State <> 3) Then
      CODIGO.ReadOnly = True
      MOTIVO.ReadOnly = True
    Else
      CODIGO.ReadOnly = False
      MOTIVO.ReadOnly = False
    End If

  Else
    CODIGO.ReadOnly = True
    MOTIVO.ReadOnly = True
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
    PRESTADOR.ReadOnly = True
    CODIGO.ReadOnly = True
    CODIGO.ReadOnly = True
    MOTIVO.ReadOnly = True
  End If

  'SMS 74435 - Marcelo Barbosa - 28/12/2006
  'Mudada a verificação da situação, pois ao tentar incluir e clicar sobre algum registro
  'acabava deixando a situação livre para editar
  'SMS 63369 [inicio]
  If CurrentQuery.State <> 3 Then
    SITUACAO.ReadOnly = True
  Else
    SITUACAO.ReadOnly = False
  End If
  'SMS 63369 [fim]
  'SMS 74435

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If ((Not CurrentQuery.FieldByName("FATURABENEFICIARIO").IsNull) Or (Not CurrentQuery.FieldByName("FATURAPRESTADOR").IsNull)) Then
    If (VisibleMode) Then
      bsShowMessage("Cartão já foi faturado. Não pode ser excluído.", "I")
    Else
      CancelDescription = "Cartão já foi faturado. Não pode ser excluído."
    End If
    CanContinue = False
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "A" Then
    bsShowMessage("Cartão 'Ativo' não pode ser excluído.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If ((Not CurrentQuery.FieldByName("FATURABENEFICIARIO").IsNull) Or (Not CurrentQuery.FieldByName("FATURAPRESTADOR").IsNull)) Then
    If (VisibleMode) Then
      bsShowMessage("Cartão já foi faturado. Não pode ser alterado.", "I")
    Else
      CancelDescription = "Cartão já foi faturado. Não pode ser alterado."
    End If
    CanContinue = False
    Exit Sub
  End If

  If (Not CurrentQuery.FieldByName("ROTINAIMP").IsNull) Then
    CODIGO.ReadOnly = True
    MOTIVO.ReadOnly = True
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = False
    'DATACANCELAMENTO.ReadOnly = False
    PRESTADOR.ReadOnly = True
    CODIGO.ReadOnly = True
    CODIGO.ReadOnly = True
    MOTIVO.ReadOnly = True
    Exit Sub
  End If

  'SMS 69087
  If (CurrentQuery.FieldByName("SITUACAO").AsString = "P") And _
     (Not CurrentQuery.FieldByName("ROTINACONVENIORECIPRENOVACAO").IsNull) And _
     CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
     'se o registro estiver pendente
     'se o registro tiver sido gerado por uma rotina de renovação
     'se não estiver cancelado
    CODIGO.ReadOnly = False
    SITUACAO.ReadOnly = False
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim INTERFACE As Object
  Dim LINHA As String
  Dim Condicao As String
  Dim qBeneficiario As Object

  Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = " AND BENEFICIARIO = " + CurrentQuery.FieldByName("BENEFICIARIO").AsString + " AND DATACANCELAMENTO IS NULL "

  If VisibleMode = True Then
    LINHA = INTERFACE.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_CONVENIORECIP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)
    If (LINHA <> "") Then
      If (VisibleMode) Then
        bsShowMessage(LINHA, "I")
      Else
        CancelDescription = LINHA
      End If
      CanContinue = False
    End If
  End If

  Set INTERFACE = Nothing

  Set qBeneficiario = NewQuery
  qBeneficiario.Add("SELECT DATAADESAO, DATACANCELAMENTO ")
  qBeneficiario.Add("  FROM SAM_BENEFICIARIO ")
  qBeneficiario.Add(" WHERE HANDLE = " + CurrentQuery.FieldByName("BENEFICIARIO").AsString)
  qBeneficiario.Active = True
  If (Not qBeneficiario.FieldByName("DATACANCELAMENTO").IsNull) Then
    If (qBeneficiario.FieldByName("DATACANCELAMENTO").AsDateTime < ServerDate) Then
      If (VisibleMode) Then
        bsShowMessage("Beneficiário cancelado. Não é possível inserir Cartões de Reciprocidade.", "I")
      Else
        CancelDescription = "Beneficiário cancelado. Não é possível inserir Cartões de Reciprocidade."
      End If
      CanContinue = False
      Set qBeneficiario = Nothing
      Exit Sub
    End If

    If (qBeneficiario.FieldByName("DATACANCELAMENTO").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) Then
      If (VisibleMode) Then
        bsShowMessage("Não é possível inserir cartões com esta Data Inicial. Beneficiário estará cancelado.", "I")
      Else
        CancelDescription = "Não é possível inserir cartões com esta Data Inicial. Beneficiário estará cancelado."
      End If
      CanContinue = False
      Set qBeneficiario = Nothing
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
      If (VisibleMode) Then
        bsShowMessage("Não é possível inserir cartões sem Data Final. Beneficiário está cancelado com data futura.", "I")
      Else
        CancelDescription = "Não é possível inserir cartões sem Data Final. Beneficiário está cancelado com data futura."
      End If
      CanContinue = False
      Set qBeneficiario = Nothing
      Exit Sub
    End If

    If ((Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And (qBeneficiario.FieldByName("DATACANCELAMENTO").AsDateTime < CurrentQuery.FieldByName("DATAFINAL").AsDateTime)) Then
      If (VisibleMode) Then
        bsShowMessage("Não é possível inserir cartões com esta Data Final. Beneficiário estará cancelado.", "I")
      Else
        CancelDescription = "Não é possível inserir cartões com esta Data Final. Beneficiário estará cancelado."
      End If
      CanContinue = False
      Set qBeneficiario = Nothing
      Exit Sub
    End If
  End If

  If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < qBeneficiario.FieldByName("DATAADESAO").AsDateTime) Then
    If (VisibleMode) Then
      bsShowMessage("Data Inicial menor que a Data de Adesão do Beneficiário.", "I")
    Else
      CancelDescription = "Data Inicial menor que a Data de Adesão do Beneficiário."
    End If
    CanContinue = False
    Set qBeneficiario = Nothing
    Exit Sub
  End If


  'Considerar apenas na inclusão do registro
  If CurrentQuery.State = 3 Then

    Set qBeneficiario = Nothing

    Dim dllBSBen017_Rotinas As Object
    Dim viCodigoRetorno As Long
    Dim vsMensagemRetorno As String

    Set dllBSBen017_Rotinas = CreateBennerObject("BSBEN017.Rotinas")

    'Validar a inclusão do registro conforme os parâmetros do contrato
    dllBSBen017_Rotinas.ValidarConvenioReciprocidade(CurrentSystem, _
                                                     CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, _
                                                     CurrentQuery.FieldByName("PRESTADOR").AsInteger, _
                                                     CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
                                                     viCodigoRetorno, _
                                                     vsMensagemRetorno)

    If viCodigoRetorno > 0 Then
      Set dllBSBen017_Rotinas = Nothing
      CanContinue = False
      If VisibleMode Then
        bsShowMessage("O registro não pode ser incluído - " + vsMensagemRetorno, "Ï")
      Else
        CancelDescription = "O registro não pode ser incluído - " + vsMensagemRetorno
      End If
    End If
    Set dllBSBen017_Rotinas = Nothing
  End If

End Sub
