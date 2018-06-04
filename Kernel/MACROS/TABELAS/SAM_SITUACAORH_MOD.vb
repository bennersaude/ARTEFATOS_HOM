'HASH: 2E447F1B77A498B2FCFB82651BC23BB6
'#Uses "*bsShowMessage"

Dim vContrato As Integer

Public Sub PLANODESTINO_OnPopup(ShowPopup As Boolean)
  Dim SelContrato As Object

  Set SelContrato = NewQuery

  SelContrato.Active = False
  SelContrato.Clear
  SelContrato.Add("SELECT CONTRATOMIGRACAO FROM SAM_SITUACAORH WHERE HANDLE = :HANDLE")
  SelContrato.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SITUACAORH").AsInteger
  SelContrato.Active = True


  vContrato = SelContrato.FieldByName("CONTRATOMIGRACAO").AsInteger
  PLANODESTINO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = " + Str(vContrato) + ")"

  Set SelContrato = Nothing
End Sub

Public Sub PLANOORIGEM_OnPopup(ShowPopup As Boolean)
  Dim SelContrato As Object

  Set SelContrato = NewQuery

  SelContrato.Active = False
  SelContrato.Clear
  SelContrato.Add("SELECT CONTRATO FROM SAM_SITUACAORH WHERE HANDLE = :HANDLE")
  SelContrato.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SITUACAORH").AsInteger
  SelContrato.Active = True


  vContrato = SelContrato.FieldByName("CONTRATO").AsInteger
  PLANOORIGEM.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = " + Str(vContrato) + ")"

  Set SelContrato = Nothing
End Sub

Public Sub MODULODESTINO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vPlano As Integer
  Dim SQL As Object
  Dim qModuloOrigem As Object

  If CurrentQuery.FieldByName("MODULOORIGEM").IsNull Then
   BsShowMessage("Informe o módulo de origem.","E")
   ShowPopup = False

   Exit Sub
  End If

  Set SQL = NewQuery
  Set SelContrato = NewQuery

  SelContrato.Active = False
  SelContrato.Clear
  SelContrato.Add("SELECT CONTRATOMIGRACAO FROM SAM_SITUACAORH WHERE HANDLE = :HANDLE")
  SelContrato.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SITUACAORH").AsInteger
  SelContrato.Active = True
  vPlano = CurrentQuery.FieldByName("PLANODESTINO").AsInteger

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_MODULO.DESCRICAO"

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE                ")
  SQL.Add("  FROM SAM_MODEQUIVALENCIA   ")
  SQL.Active = True

  If Not SQL.EOF Then

    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT MODULO                ")
    SQL.Add("  FROM SAM_CONTRATO_MOD      ")
    SQL.Add(" WHERE HANDLE = :HANDLE      ")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
    SQL.Active = True



    vCriterio = " SAM_CONTRATO_MOD.CONTRATO = " + SelContrato.FieldByName("CONTRATOMIGRACAO").AsString + _
                " AND SAM_CONTRATO_MOD.PLANO = " + CStr(vPlano) + _
                " AND SAM_CONTRATO_MOD.MODULO IN (SELECT MODULO FROM SAM_MODEQUIVALENCIA_MOD " +  _
                "                                  WHERE MODEQUIVALENCIA IN (SELECT MODEQUIVALENCIA " + _
                "                                                               FROM SAM_MODEQUIVALENCIA_MOD WHERE MODULO = "+ SQL.FieldByName("MODULO").AsString + _
                "                                                            ))"
    vCampos = "Módulo|Plano"

    vHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]", vColunas, 1, vCampos, vCriterio, "Módulos", True, MODULODESTINO.Text)

  Else

    'sms 74054
    Set qModuloOrigem = NewQuery
    qModuloOrigem.Active = False
    qModuloOrigem.Clear
    qModuloOrigem.Add("SELECT CM.OBRIGATORIO, CM.REGISTROMS, MS.NOVAREGULAMENTACAO  ")
    qModuloOrigem.Add("  FROM SAM_MODULO M, SAM_CONTRATO_MOD CM                          ")
    qModuloOrigem.Add("       LEFT JOIN SAM_REGISTROMS MS ON (CM.REGISTROMS = MS.HANDLE) ")
    qModuloOrigem.Add(" WHERE CM.MODULO = M.HANDLE        ")
    qModuloOrigem.Add("       AND CM.HANDLE = :MODORIGEM  ")
    qModuloOrigem.ParamByName("MODORIGEM").AsInteger = CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
    qModuloOrigem.Active = True

    vCriterio = " SAM_CONTRATO_MOD.CONTRATO = " + SelContrato.FieldByName("CONTRATOMIGRACAO").AsString + _
                " AND SAM_CONTRATO_MOD.PLANO = " + CStr(vPlano) + _
                " AND SAM_CONTRATO_MOD.OBRIGATORIO = '" + qModuloOrigem.FieldByName("OBRIGATORIO").AsString + "'"

    If Not qModuloOrigem.FieldByName("REGISTROMS").IsNull Then
      vCriterio = vCriterio + " AND (SAM_CONTRATO_MOD.REGISTROMS IS NOT NULL) "
      If qModuloOrigem.FieldByName("NOVAREGULAMENTACAO").AsString = "S" Then
        vCriterio = vCriterio + " AND SAM_REGISTROMS.NOVAREGULAMENTACAO = 'S' "
      End If
    End If

    vCampos = "Módulo|Plano"

    If qModuloOrigem.FieldByName("NOVAREGULAMENTACAO").AsString = "S" Then
      vHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]|SAM_REGISTROMS[SAM_REGISTROMS.HANDLE = SAM_CONTRATO_MOD.REGISTROMS]", vColunas, 1, vCampos, vCriterio, "Módulos", True, MODULODESTINO.Text)
    Else
      vHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]", vColunas, 1, vCampos, vCriterio, "Módulos", True, MODULODESTINO.Text)
    End If

    Set qModuloOrigem = Nothing

  End If


  If vHandle <>0 Then
    CurrentQuery.FieldByName("MODULODESTINO").Value = vHandle
  End If
  Set Interface = Nothing

  Set SelContrato = Nothing
  Set SQL = Nothing
End Sub

Public Sub MODULOORIGEM_OnPopup(ShowPopup As Boolean)

  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vPlano As Integer

  Set SelContrato = NewQuery

  SelContrato.Active = False
  SelContrato.Clear
  SelContrato.Add("SELECT CONTRATO FROM SAM_SITUACAORH WHERE HANDLE = :HANDLE")
  SelContrato.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SITUACAORH").AsInteger
  SelContrato.Active = True
  vPlano = CurrentQuery.FieldByName("PLANOORIGEM").AsInteger

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_MODULO.DESCRICAO"

  vCriterio = " SAM_CONTRATO_MOD.CONTRATO = " + SelContrato.FieldByName("CONTRATO").AsString + _
              " AND SAM_CONTRATO_MOD.PLANO = " + CStr(vPlano) + " AND SAM_CONTRATO_MOD.HANDLE NOT IN (SELECT MODULOORIGEM FROM SAM_SITUACAORH_MOD WHERE SITUACAORH = " + Str(CurrentQuery.FieldByName("SITUACAORH").AsInteger) + " ) "

  vCampos = "Módulo|Plano"

  vHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]", vColunas, 1, vCampos, vCriterio, "Módulos", True, MODULOORIGEM.Text)

  If vHandle <>0 Then
    CurrentQuery.FieldByName("MODULOORIGEM").Value = vHandle
  End If
  Set Interface = Nothing

  Set SelContrato = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'sms 74054
  Dim qModulo1 As Object
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE                ")
  SQL.Add("  FROM SAM_MODEQUIVALENCIA   ")
  SQL.Active = True

  If SQL.EOF Then

    Dim qModOrigem As Object
    Dim qModDestino As Object

    Set qModOrigem = NewQuery
    Set qModDestino = NewQuery

    qModOrigem.Active = False
    qModOrigem.Clear
    qModOrigem.Add("SELECT M.TIPOMODULO, M.TIPOCOBERTURA, CM.OBRIGATORIO, CM.REGISTROMS, MS.NOVAREGULAMENTACAO  ")
    qModOrigem.Add("  FROM SAM_MODULO M, SAM_CONTRATO_MOD CM                          ")
    qModOrigem.Add("       LEFT JOIN SAM_REGISTROMS MS ON (CM.REGISTROMS = MS.HANDLE) ")
    qModOrigem.Add(" WHERE CM.MODULO = M.HANDLE        ")
    qModOrigem.Add("       AND CM.HANDLE = :MODORIGEM  ")
    qModOrigem.ParamByName("MODORIGEM").AsInteger = CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
    qModOrigem.Active = True

    qModDestino.Active = False
    qModDestino.Clear
    qModDestino.Add("SELECT M.TIPOMODULO, M.TIPOCOBERTURA, CM.OBRIGATORIO, CM.REGISTROMS, MS.NOVAREGULAMENTACAO  ")
    qModDestino.Add("  FROM SAM_MODULO M, SAM_CONTRATO_MOD CM                          ")
    qModDestino.Add("       LEFT JOIN SAM_REGISTROMS MS ON (CM.REGISTROMS = MS.HANDLE) ")
    qModDestino.Add(" WHERE CM.MODULO = M.HANDLE        ")
    qModDestino.Add("       AND CM.HANDLE = :MODDESTINO  ")
    qModDestino.ParamByName("MODDESTINO").AsInteger = CurrentQuery.FieldByName("MODULODESTINO").AsInteger
    qModDestino.Active = True

    If (qModOrigem.FieldByName("OBRIGATORIO").AsString = "S") And (qModDestino.FieldByName("OBRIGATORIO").AsString <> "S") Then
      BsShowMessage("Para módulo de origem obrigatório deve ser cadastrado módulo destino obrigatório","E")
      CanContinue = False
      Exit Sub
    End If
    If (qModOrigem.FieldByName("OBRIGATORIO").AsString = "N") And (qModDestino.FieldByName("OBRIGATORIO").AsString <> "N") Then
      BsShowMessage("Para módulo de origem opcional deve ser cadastrado módulo destino opcional","E")
      CanContinue = False
      Exit Sub
    End If

    If Not qModOrigem.FieldByName("REGISTROMS").IsNull Then
      If qModDestino.FieldByName("REGISTROMS").IsNull Then
        BsShowMessage("Módulo destino deve ter Registro Ministério da Saúde","E")
        CanContinue = False
        Exit Sub
      End If

      If (qModOrigem.FieldByName("NOVAREGULAMENTACAO").AsString = "S") Then
        If (qModDestino.FieldByName("NOVAREGULAMENTACAO").AsString <> "S") Then
          BsShowMessage("Módulo destino deve ter o parâmetro 'Nova Regulamentação' marcado","E")
          CanContinue = False
          Exit Sub
        End If
      End If
    End If
    Set qModOrigem = Nothing
    Set qModDestino = Nothing

  End If

  Set SQL = Nothing


End Sub

