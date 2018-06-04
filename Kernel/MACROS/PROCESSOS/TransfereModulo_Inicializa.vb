'HASH: 75911A38F27C2E72A055FF5B927553AE

Public Sub Main
	On Error GoTo erro

	Dim BsBen002Dll As Object


	Dim psRotBenef As String
	Dim psModOrigem As String
	Dim psMensagem As String
	Dim plModOrigem As Long
	Dim plPlanoOrigem As Long
	Dim psPlanoOrigem As String
	Dim psMotCanc As String
	Dim plMotCanc As Long
	Dim plDiasCarencia As Long
	Dim pdDataAdesao As Date
	Dim psCodigoTabPrc As String
	Dim psCarencia As String
	Dim psCriterios As String
	Dim vlHModBeneficiario As Long



	Set BsBen002Dll = CreateBennerObject( "BSBEN002.Modulo" )

	psRotBenef = CStr( ServiceVar("psRotBenef") )
	psModOrigem = CStr( ServiceVar("psModOrigem") )
	psPlanoOrigem = CStr( ServiceVar("psPlanoOrigem"))
	psMotCanc = CStr( ServiceVar("psMotCanc") )
	plDiasCarencia = CLng( ServiceVar("plDiasCarencia") )
	pdDataAdesao = CDate( ServiceVar("pdDataAdesao") )
	psCodigoTabPrc = CStr(ServiceVar("psCodigoTabPrc"))
	psCarencia = CStr( ServiceVar("psCarencia") )
	psCriterios = CStr(ServiceVar("psCriterios") )
	psMensagem = CStr( ServiceVar( "psMensagem") )
	vlHModBeneficiario = CLng( SessionVar("HMODBENEFICIARIO") )

	psCarencia = BsBen002Dll.InitInterface( CurrentSystem, _
											  vlHModBeneficiario,  _
     										  psRotBenef, _
											  plModOrigem, _
             								  plPlanoOrigem, _
											  plMotCanc, _
											  plDiasCarencia, _
											  pdDataAdesao )

	'ServiceResult = "retorno"
	Dim SQL As Object
	Set SQL =  NewQuery

	SQL.Clear
	SQL.Add("SELECT DESCRICAO FROM SAM_MODULO  WHERE HANDLE = (:HANDLE)")
	SQL.ParamByName("HANDLE").AsInteger = plModOrigem
    SQL.Active = True

	psModOrigem = SQL.FieldByName("DESCRICAO").AsString

	ServiceVar("psRotBenef") = CStr( psRotBenef )
	ServiceVar("psModOrigem") = psModOrigem

    SQL.Active = False
	SQL.Clear

	SQL.Add("SELECT DESCRICAO FROM SAM_PLANO WHERE HANDLE = (:HANDLE)")
	SQL.ParamByName("HANDLE").AsInteger = plPlanoOrigem
    SQL.Active = True

    psPlanoOrigem = SQL.FieldByName("DESCRICAO").AsString

	ServiceVar("psPlanoOrigem") = psPlanoOrigem

	SQL.Active = False
	SQL.Clear

	SQL.Add("SELECT DESCRICAO FROM  SAM_MOTIVOCANCELAMENTO WHERE HANDLE = (:HANDLE)")
	SQL.ParamByName("HANDLE").AsInteger = plMotCanc

	SQL.Active = True
	ServiceVar("plMotCanc") = plMotCanc
	ServiceVar("psMotCanc") = CStr(SQL.FieldByName("DESCRICAO").AsString)


	ServiceVar("plDiasCarencia") = CLng( plDiasCarencia )
	ServiceVar("pdDataAdesao") = ServerDate

	SQL.Active = False
	SQL.Clear

	SQL.Add("SELECT CODIGOTABELAPRC FROM SAM_BENEFICIARIO_MOD WHERE HANDLE = (:HANDLE)")
	SQL.ParamByName("HANDLE").AsInteger =vlHModBeneficiario
	SQL.Active = True

	psCodigoTabPrc = SQL.FieldByName("CODIGOTABELAPRC").AsString

	ServiceVar("psCodigoTabPrc") = CStr( psCodigoTabPrc )

	ServiceVar("psCarencia") = CStr( psCarencia )


    SQL.Active = False
    SQL.Clear

	SQL.Add("SELECT B.CONTRATO, B.FAMILIA								   ")
    SQL.Add("  FROM SAM_BENEFICIARIO B                                     ")
    SQL.Add("  JOIN SAM_BENEFICIARIO_MOD BM ON (B.HANDLE = BM.BENEFICIARIO)")
    SQL.Add(" WHERE BM.HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = vlHModBeneficiario

    SQL.Active = True

    Dim vsAux As String
    Dim vsFamilia As Long


    vsAux = CStr(SQL.FieldByName("CONTRATO").AsInteger)
    vsFamilia = SQL.FieldByName("FAMILIA").AsInteger

    psCriterios =  "SAM_CONTRATO_MOD.CONTRATO = " + vsAux + _
   	    	       " AND ((SAM_CONTRATO_MOD.MODULO <> " + CStr(plModOrigem) + ")" + _
  				   "  OR  (SAM_CONTRATO_MOD.MODULO = " + CStr(plModOrigem) + _
                   " AND SAM_CONTRATO_MOD.PLANO <> " + CStr(plPlanoOrigem) + "))"

    SQL.Active = False
    SQL.Clear

    SQL.Add("SELECT OBRIGATORIO , TIPOMODULO, COM.PLANO")
    SQL.Add("  FROM SAM_CONTRATO_MOD COM,SAM_MODULO MOD")
    SQL.Add(" WHERE COM.MODULO=MOD.HANDLE              ")
    SQL.Add("   AND COM.MODULO = " + CStr(plModOrigem) +    " And COM.CONTRATO = " +vsAux)

	SQL.Active = True

   vsAux = SQL.FieldByName("OBRIGATORIO").AsString



 	psCriterios = psCriterios + " AND SAM_CONTRATO_MOD.OBRIGATORIO = '" + vsAux +  "'"

    If Not (SQL.FieldByName("TIPOMODULO").AsString = "S" ) Then
      psCriterios = psCriterios + " AND SAM_MODULO.TIPOMODULO <> 'S'    "
    End If
  	psCriterios = psCriterios + " AND DATACANCELAMENTO IS NULL "

    SQL.Active = False
    SQL.Clear

    SQL.Add("SELECT TIPOMODULO,     ")
    SQL.Add("       TIPOCOBERTURA   ")
    SQL.Add("  FROM SAM_MODULO      ")
    SQL.Add(" WHERE HANDLE = :MODULO")
    SQL.ParamByName("MODULO").AsInteger = plModOrigem
    SQL.Active = True


  	psCriterios = psCriterios +  " AND SAM_MODULO.TIPOMODULO = '" + SQL.FieldByName("TIPOMODULO").AsString + "'"
  	psCriterios = psCriterios +  " AND SAM_MODULO.TIPOCOBERTURA = '"  + SQL.FieldByName("TIPOCOBERTURA").AsString  + "'"

	SQL.Active = False
	SQL.Clear

    SQL.Add("SELECT A.TITULARRESPONSAVEL,")
    SQL.Add("A.TABRESPONSAVEL,")
    SQL.Add("C.DATAINCLUSAOPREVIDENCIA")
	SQL.Add("FROM SAM_FAMILIA A,")
    SQL.Add("SAM_BENEFICIARIO B, SAM_MATRICULA C")
 	SQL.Add("WHERE A.Handle = B.FAMILIA")
   	SQL.Add("And B.MATRICULA = C.Handle")
   	SQL.Add("And B.Handle = A.TITULARRESPONSAVEL")
   	SQL.Add("And A.Handle = :HFAMILIA")
   	SQL.Add("And C.DATAINCLUSAOPREVIDENCIA <= :DATAADESAO")

   	SQL.ParamByName("HFAMILIA").AsInteger = vsFamilia
    SQL.ParamByName("DATAADESAO").AsDateTime = ServerDate

    SQL.Active = True

 	If  SQL.FieldByName("DATAINCLUSAOPREVIDENCIA").IsNull Then
      psCriterios = psCriterios +   " AND SAM_CONTRATO_MOD.VERIFICARCONTRIBPREVIDENCIA =  'N' "
	End If

	ServiceVar("psCriterios") = psCriterios

	Set SQL = Nothing
	Set BsBen002Dll = Nothing
	Exit Sub

	erro:
	  psMensagem =  Err.Description
      ServiceVar( "psMensagem" ) = psMensagem

End Sub
