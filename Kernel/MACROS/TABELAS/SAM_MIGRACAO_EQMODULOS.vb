'HASH: A59AFEB09AB778A29DFE42B0AE469AA6
'Macro: SAM_MIGRACAO_EQMODULOs
'#Uses "*bsShowMessage"


Dim vPlano As Integer
Dim vContrato As Integer

Public Sub MODULODESTINO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vPlano As Integer

  Dim qModuloOrigem As Object 'sms 74054

  If CurrentQuery.FieldByName("MODULOORIGEM").IsNull Then
   bsShowMessage("Informar o módulo de origem.", "I")
   ShowPopup = False

   Exit Sub
  End If



  vPlano = CurrentQuery.FieldByName("PLANODESTINO").AsInteger

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_MODULO.DESCRICAO"


  Dim moduloMigracaoBLL As CSBusinessComponent

  Set moduloMigracaoBLL = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.Modulo.ModuloMigracaoBLL, Benner.Saude.Beneficiarios.Business")

	moduloMigracaoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("CONTRATODESTINO").AsInteger)
	moduloMigracaoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("MODULOORIGEM").AsInteger)
	moduloMigracaoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PLANODESTINO").AsInteger)
	moduloMigracaoBLL.AddParameter(pdtString, "SAM_CONTRATO_MOD")

	vCriterio = moduloMigracaoBLL.Execute("CarregarCriterioSelecaoModuloOrigem")

	Set moduloMigracaoBLL = Nothing

    vCampos = "Módulo|Plano"

    vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]", vColunas, 1, vCampos, vCriterio, "Módulos", True, "")


    Set qModuloOrigem = Nothing 'sms 73847


  If vHandle <>0 Then
    CurrentQuery.FieldByName("MODULODESTINO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub MODULOORIGEM_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vPlano As Integer

  vPlano = CurrentQuery.FieldByName("PLANOORIGEM").AsInteger

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_MODULO.DESCRICAO"

  vCriterio = " SAM_CONTRATO_MOD.CONTRATO = " + CurrentQuery.FieldByName("CONTRATOORIGEM").AsString + _
              " AND SAM_CONTRATO_MOD.PLANO = " + CStr(vPlano)
  vCampos = "Módulo|Plano"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_MODULO.HANDLE = SAM_CONTRATO_MOD.MODULO]", vColunas, 1, vCampos, vCriterio, "Módulos", True, "")

  If vHandle <>0 Then
    CurrentQuery.FieldByName("MODULOORIGEM").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()

	Dim SQL As Object

	Set SQL = NewQuery

  	SQL.Active = False
  	SQL.Clear
  	SQL.Add("SELECT HANDLE                ")
  	SQL.Add("  FROM SAM_MODEQUIVALENCIA   ")
  	SQL.Active = True

	If Not SQL.EOF Then

		If WebMode Then
	 		MODULODESTINO.WebLocalWhere = " A.CONTRATO = @CAMPO(CONTRATODESTINO)" + _
 	              					  	" AND A.PLANO = @CAMPO(PLANODESTINO)" + _
                  					  	" AND A.MODULO IN (SELECT MODULO FROM SAM_MODEQUIVALENCIA_MOD " +  _
                  					  	"                                  WHERE MODEQUIVALENCIA IN (SELECT MODEQUIVALENCIA " + _
                  					  	"                                                               FROM SAM_MODEQUIVALENCIA_MOD WHERE MODULO IN" + _
                  					  	"															(SELECT MODULO FROM SAM_CONTRATO_MOD WHERE HANDLE = @CAMPO(MODULOORIGEM)) " + _
                  					  	"                                                            ))"
		End If


	Else


	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If WebMode Then
  	PLANODESTINO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATODESTINO))"
  	PLANOORIGEM.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATOORIGEM))"
  ElseIf VisibleMode Then
  	PLANODESTINO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATODESTINO)"
  	PLANOORIGEM.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATOORIGEM)"
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
  	PLANODESTINO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATODESTINO))"
  	PLANOORIGEM.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATOORIGEM))"
  ElseIf VisibleMode Then
  	PLANODESTINO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATODESTINO)"
  	PLANOORIGEM.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATOORIGEM)"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'sms 50287
  Dim qModulo1 As Object
  Dim SQL As Object
  Set SQL = NewQuery
  Set qModulo1 = NewQuery
  qModulo1.Active = False
  qModulo1.Clear
  qModulo1.Add("SELECT HANDLE")
  qModulo1.Add("  FROM SAM_MIGRACAO_EQMODULOS")
  qModulo1.Add(" WHERE CONTRATOORIGEM = :CONTRATOORIGEM AND CONTRATODESTINO = :CONTRATODESTINO")
  qModulo1.Add("       AND MODULOORIGEM = :MODULOORIGEM AND HANDLE <> :HANDLE")

  qModulo1.ParamByName("CONTRATOORIGEM").AsInteger = CurrentQuery.FieldByName("CONTRATOORIGEM").AsInteger
  qModulo1.ParamByName("CONTRATODESTINO").AsInteger = CurrentQuery.FieldByName("CONTRATODESTINO").AsInteger
  qModulo1.ParamByName("MODULOORIGEM").AsInteger = CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
  qModulo1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qModulo1.Active = True

  If Not qModulo1.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Equivalência existente para o módulo origem informado", "E")
    CanContinue = False
    Set qModulo1 = Nothing
    Exit Sub
  End If
  Set qModulo1 = Nothing

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
      bsShowMessage("Para módulo de origem obrigatório deve ser cadastrado módulo destino obrigatório", "E")
      CanContinue = False
      Exit Sub
    End If
    If (qModOrigem.FieldByName("OBRIGATORIO").AsString = "N") And (qModDestino.FieldByName("OBRIGATORIO").AsString <> "N") Then
      bsShowMessage("Para módulo de origem opcional deve ser cadastrado módulo destino opcional", "E")
      CanContinue = False
      Exit Sub
    End If

    If Not qModOrigem.FieldByName("REGISTROMS").IsNull Then
      If qModDestino.FieldByName("REGISTROMS").IsNull Then
        bsShowMessage("Módulo destino deve ter Registro Ministério da Saúde", "E")
        CanContinue = False
        Exit Sub
      End If

      If (qModOrigem.FieldByName("NOVAREGULAMENTACAO").AsString = "S") Then
        If (qModDestino.FieldByName("NOVAREGULAMENTACAO").AsString <> "S") Then
          bsShowMessage("Módulo destino deve ter o parâmetro 'Nova Regulamentação' marcado", "E")
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

