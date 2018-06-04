'HASH: 8CB4009AC2082ADB74E29E50C6A42969
'#Uses "*bsShowMessage"

Public Sub RESPONSAVELNOVO_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  Dim Interface As Object
  Dim viHBeneficiario As Long
  Dim vsColunas As String
  Dim vsCampos As String
  Dim vsCriterio As String

  vsColunas  = "Beneficiário|Nome|Data de cancelamento"
  vsCampos   = "BENEFICIARIO|NOME|DATACANCELAMENTO"
  vsCriterio = " FAMILIA = " + CurrentQuery.FieldByName("FAMILIA").AsString + _
               " AND (DATACANCELAMENTO IS NULL OR DATACANCELAMENTO > " + SQLDate(ServerDate) + ")" + _
               " AND HANDLE <> " + CurrentQuery.FieldByName("RESPONSAVELATUAL").AsString

  Set Interface = CreateBennerObject("PROCURA.Procurar")
  viHBeneficiario = Interface.Exec(CurrentSystem, _
                                   "SAM_BENEFICIARIO", _
                                   vsCampos, _
                                   1, _
                                   vsColunas, _
                                   vsCriterio, _
                                   "Selecione o novo beneficiário titular", _
                                   True, _
                                   "")

  If viHBeneficiario > 0 Then
    CurrentQuery.FieldByName("RESPONSAVELNOVO").AsInteger = viHBeneficiario
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface  As Object
  Dim viRetorno  As Integer
  Dim vsMensagem As String

  Set Interface = CreateBennerObject("BSBEN019.TrocaTitular")

  viRetorno = Interface.Exec(CurrentSystem, _
                             CurrentQuery.FieldByName("FAMILIA").AsInteger, _
                             CurrentQuery.FieldByName("RESPONSAVELNOVO").AsInteger, _
                             CurrentQuery.FieldByName("NOVOTPDEPRESPONSAVELATUAL").AsInteger, _
                             CurrentQuery.FieldByName("TPDEPRESPONSAVELNOVO").AsInteger, _
                             vsMensagem)

  If viRetorno = 1 Then
    bsShowMessage(vsMensagem, "E")
    CanContinue = False
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_NewRecord()
  Dim viHFamilia As Long

  If WebMode Then
    CurrentQuery.FieldByName("FAMILIA").AsInteger = RecordHandleOfTable("SAM_FAMILIA")
    viHFamilia = RecordHandleOfTable("SAM_FAMILIA")
  Else
    CurrentQuery.FieldByName("FAMILIA").AsInteger = CLng(SessionVar("HFAMILIA_TROCATITULAR"))
    viHFamilia = CLng(SessionVar("HFAMILIA_TROCATITULAR"))
  End If

  Dim SQL As Object
  Dim vsCondicaoTpDep As String
  Set SQL = NewQuery

  SQL.Add("SELECT  FAM.TITULARRESPONSAVEL,")
  SQL.Add("        FAM.CONTRATO,")
  SQL.Add("       CONT.PERMITIRDEPENDENTETITULAR")
  SQL.Add("FROM SAM_FAMILIA FAM")
  SQL.Add("JOIN SAM_CONTRATO CONT ON CONT.HANDLE = FAM.CONTRATO")
  SQL.Add("WHERE FAM.HANDLE = :HFAMILIA")
  SQL.ParamByName("HFAMILIA").AsInteger = viHFamilia
  SQL.Active = True

  CurrentQuery.FieldByName("CONTRATO").AsInteger = SQL.FieldByName("CONTRATO").AsInteger
  If Not SQL.FieldByName("TITULARRESPONSAVEL").IsNull Then
    CurrentQuery.FieldByName("RESPONSAVELATUAL").AsInteger = SQL.FieldByName("TITULARRESPONSAVEL").AsInteger
  End If

  If SQL.FieldByName("PERMITIRDEPENDENTETITULAR").AsString = "S" Then
    vsCondicaoTpDep = "TIPODEPENDENTE IN (SELECT TPDEP.HANDLE FROM SAM_TIPODEPENDENTE TPDEP WHERE TPDEP.GRUPODEPENDENTE in ('T', 'D'))"
  Else
    vsCondicaoTpDep = "TIPODEPENDENTE IN (SELECT TPDEP.HANDLE FROM SAM_TIPODEPENDENTE TPDEP WHERE TPDEP.GRUPODEPENDENTE = 'T')"
  End If

  If WebMode Then
    NOVOTPDEPRESPONSAVELATUAL.WebLocalWhere = " A.TIPODEPENDENTE IN (SELECT TPDEP.HANDLE FROM SAM_TIPODEPENDENTE TPDEP WHERE TPDEP.GRUPODEPENDENTE = 'D' OR TPDEP.GRUPODEPENDENTE = 'A') "
    TPDEPRESPONSAVELNOVO.WebLocalWhere      = " A." + vsCondicaoTpDep
    RESPONSAVELNOVO.WebLocalWhere           = " A.FAMILIA = " + CurrentQuery.FieldByName("FAMILIA").AsString + _
                                              " AND (A.DATACANCELAMENTO IS NULL OR A.DATACANCELAMENTO > " + SQLDate(ServerDate) + ")" + _
                                              " AND A.HANDLE <> " + CurrentQuery.FieldByName("RESPONSAVELATUAL").AsString
  Else
    NOVOTPDEPRESPONSAVELATUAL.LocalWhere = " SAM_CONTRATO_TPDEP.TIPODEPENDENTE IN (SELECT TPDEP.HANDLE FROM SAM_TIPODEPENDENTE TPDEP WHERE TPDEP.GRUPODEPENDENTE = 'D' OR TPDEP.GRUPODEPENDENTE = 'A') "
    TPDEPRESPONSAVELNOVO.LocalWhere      = " SAM_CONTRATO_TPDEP." + vsCondicaoTpDep
    RESPONSAVELNOVO.LocalWhere           = " SAM_BENEFICIARIO.FAMILIA = " + CurrentQuery.FieldByName("FAMILIA").AsString + _
                                           " AND (SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL OR SAM_BENEFICIARIO.DATACANCELAMENTO > " + SQLDate(ServerDate) + ")" + _
                                           " AND SAM_BENEFICIARIO.HANDLE <> " + CurrentQuery.FieldByName("RESPONSAVELATUAL").AsString
  End If

  Set SQL = Nothing
End Sub
