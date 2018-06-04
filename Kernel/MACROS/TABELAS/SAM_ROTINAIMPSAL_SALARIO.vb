'HASH: AB7B8331506397FA1BAE19552207681F
'Macro da tabela: SAM_ROTINAIMPSAL_SALARIO

Option Explicit

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  Dim Interface     As Object
  Dim viHandle      As Long
  Dim VsTabelas     As String
  Dim vsCampos      As String
  Dim vsColunas     As String
  Dim vsCriterio    As String
  Dim vsNvl         As String
  Dim vsInfinito    As String
  Dim vsCompetencia As String

  VsTabelas  = "SAM_BENEFICIARIO|SAM_CONTRATO[SAM_CONTRATO.HANDLE = SAM_BENEFICIARIO.CONTRATO]"
  vsColunas  = "SAM_CONTRATO.CONTRATO|SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_BENEFICIARIO.BENEFICIARIO|SAM_BENEFICIARIO.NOME|SAM_BENEFICIARIO.DATACANCELAMENTO"
  vsCampos   = "Contrato|Matrícula funcional|Código do beneficiário|Nome|Data de cancelamento"

  Dim qRotinaImpSal As Object
  Set qRotinaImpSal = NewQuery

  qRotinaImpSal.Clear
  qRotinaImpSal.Add("SELECT COMPETENCIA,    ")
  qRotinaImpSal.Add("       TABFILTRO,      ")
  qRotinaImpSal.Add("       GRUPOCONTRATO,  ")
  qRotinaImpSal.Add("       CONTRATOINICIAL,")
  qRotinaImpSal.Add("       CONTRATOFINAL   ")
  qRotinaImpSal.Add("  FROM SAM_ROTINAIMPSAL")
  qRotinaImpSal.Add(" WHERE HANDLE = :HANDLE")
  qRotinaImpSal.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAIMPSAL").AsInteger
  qRotinaImpSal.Active = True


  vsInfinito    = SQLAddYear(SQLDate(ServerDate),"200")
  vsCompetencia = SQLDate(qRotinaImpSal.FieldByName("COMPETENCIA").AsDateTime)
  If (InStr(SQLServer, "MSSQL") > 0) Then
    vsNvl         = "ISNULL"
  Else
    If (InStr(SQLServer,"ORACLE") > 0) Or (InStr(SQLServer,"CACHE") > 0) Then
      vsNvl         = "NVL"
    Else
      vsNvl         = "COALESCE"
    End If
  End If

  

  'Selecionar beneficiários ativos (data de cancelamento nula ou maior/igual a competência da rotina)
  vsCriterio = ""
  vsCriterio = vsCriterio  + vsNvl + "(SAM_BENEFICIARIO.DATACANCELAMENTO, " + vsInfinito + ") >= " + vsCompetencia

  'Se a rotina possui um filtro, selecionar apenas os beneficiários dos contratos do filtro.
  If (qRotinaImpSal.FieldByName("TABFILTRO").AsInteger = 1) Then
    vsCriterio = vsCriterio  + " AND SAM_CONTRATO.GRUPOCONTRATO = " + qRotinaImpSal.FieldByName("GRUPOCONTRATO").AsString
    If ((Not qRotinaImpSal.FieldByName("CONTRATOINICIAL").IsNull) And (Not qRotinaImpSal.FieldByName("CONTRATOFINAL").IsNull)) Then
      Dim qContrato         As Object
      Dim vsContratoInicial As String
      Dim vsContratofinal   As String

      Set qContrato = NewQuery

      qContrato.Clear
      qContrato.Add("SELECT CONTRATO        ")
      qContrato.Add("  FROM SAM_CONTRATO    ")
      qContrato.Add(" WHERE HANDLE = :HANDLE")

      'Buscar o código do contrato inicial do filtro.
      qContrato.ParamByName("HANDLE").AsInteger = qRotinaImpSal.FieldByName("CONTRATOINICIAL").AsInteger
      qContrato.Active = True
      vsContratoInicial = qContrato.FieldByName("CONTRATO").AsString

      'Buscar o código do contrato final do filtro.
      qContrato.Active = False
      qContrato.ParamByName("HANDLE").AsInteger = qRotinaImpSal.FieldByName("CONTRATOFINAL").AsInteger
      qContrato.Active = True
      vsContratofinal = qContrato.FieldByName("CONTRATO").AsString
      Set qContrato = Nothing

      vsCriterio = vsCriterio  + " AND SAM_CONTRATO.CONTRATO BETWEEN " + vsContratoInicial + " AND " + vsContratofinal
    End If
  End If

  vsCriterio = vsCriterio + " AND SAM_BENEFICIARIO.MATRICULAFUNCIONAL IS NOT NULL"
  vsCriterio = vsCriterio + " AND SAM_BENEFICIARIO.EHTITULAR = 'S'"

  Set Interface = CreateBennerObject("PROCURA.Procurar")
  viHandle = Interface.Exec(CurrentSystem, VsTabelas, vsColunas, 2, vsCampos, vsCriterio, "Beneficiários", False, BENEFICIARIO.Text)

  If (viHandle > 0) Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = viHandle

    'Atualizar os campos MATRICULAFUNCIONAL e CONTRATO no registro.
    Dim qDadosBeneficiario As Object
    Set qDadosBeneficiario = NewQuery

    qDadosBeneficiario.Clear
    qDadosBeneficiario.Add("SELECT A.MATRICULAFUNCIONAL,")
    qDadosBeneficiario.Add("       B.CONTRATO           ")
    qDadosBeneficiario.Add("  FROM SAM_BENEFICIARIO A,  ")
    qDadosBeneficiario.Add("       SAM_CONTRATO     B   ")
    qDadosBeneficiario.Add(" WHERE A.HANDLE = :HANDLE   ")
    qDadosBeneficiario.Add("   AND B.HANDLE = A.CONTRATO")
    qDadosBeneficiario.ParamByName("HANDLE").AsInteger = viHandle
    qDadosBeneficiario.Active = True

    CurrentQuery.FieldByName("MATRICULAFUNCIONAL").AsString  = qDadosBeneficiario.FieldByName("MATRICULAFUNCIONAL").AsString
    CurrentQuery.FieldByName("CONTRATO"          ).AsInteger = qDadosBeneficiario.FieldByName("CONTRATO").AsInteger
    Set qDadosBeneficiario = Nothing
  End If
  Set Interface     = Nothing
  Set qRotinaImpSal = Nothing
End Sub

Public Sub PESSOA_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  Dim Interface  As Object
  Dim viHandle   As Long
  Dim VsTabelas  As String
  Dim vsCampos   As String
  Dim vsColunas  As String
  Dim vsCriterio As String

  VsTabelas  = "SFN_PESSOA"
  vsColunas  = "MATRICULAFUNCIONAL|CNPJCPF|NOME"
  vsCampos   = "Matrícula funcional|CNPJ/CPF|Nome"
  vsCriterio = "MATRICULAFUNCIONAL IS NOT NULL AND TABFISICAJURIDICA = 1"

  Set Interface = CreateBennerObject("PROCURA.Procurar")
  viHandle = Interface.Exec(CurrentSystem, VsTabelas, vsColunas, 1, vsCampos, vsCriterio, "Pessoas", False, PESSOA.Text)

  If (viHandle > 0) Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PESSOA").AsInteger = viHandle

    'Atualizar o campo MATRICULAFUNCIONAL do registro.
    Dim qDadosPessoa As Object
    Set qDadosPessoa = NewQuery

    qDadosPessoa.Clear
    qDadosPessoa.Add("SELECT MATRICULAFUNCIONAL")
    qDadosPessoa.Add("  FROM SFN_PESSOA        ")
    qDadosPessoa.Add(" WHERE HANDLE = :HANDLE  ")
    qDadosPessoa.ParamByName("HANDLE").AsInteger = viHandle
    qDadosPessoa.Active = True

    CurrentQuery.FieldByName("MATRICULAFUNCIONAL").AsString = qDadosPessoa.FieldByName("MATRICULAFUNCIONAL").AsString
    Set qDadosPessoa = Nothing
  End If
  Set Interface     = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("TABVINCULO").IsNull) Then
    CurrentQuery.FieldByName("BENEFICIARIO").Clear
    CurrentQuery.FieldByName("PESSOA").Clear
  Else
    If (CurrentQuery.FieldByName("TABVINCULO").AsInteger = 1) Then
      CurrentQuery.FieldByName("PESSOA").Clear
    Else
      CurrentQuery.FieldByName("BENEFICIARIO").Clear
    End If
  End If
End Sub
