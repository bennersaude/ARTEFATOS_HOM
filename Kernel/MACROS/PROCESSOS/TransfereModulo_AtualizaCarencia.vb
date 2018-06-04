'HASH: 60893515B941EB6B1FB45A1FFB39BB9B

Public Sub Main
	On Error GoTo erro

	Dim pHContratoMod As Long
	Dim psDataAdesao As String
	Dim psCarencia As String
	Dim psCarenciaDestino As String
	Dim psPlanoDestino As String
	Dim plPlanoDestino As Long
	Dim psMensagem As String


    Dim vlHModBeneficiario As Long

	vlHModBeneficiario = CLng( SessionVar("HMODBENEFICIARIO") )

	pHContratoMod =CLng( ServiceVar("pHContratoMod") )
	psDataAdesao = CStr( ServiceVar("psDataAdesao") )
	psCarencia = CStr( ServiceVar("psCarencia") )
	psCarenciaDestino = CStr( ServiceVar("psCarenciaDestino") )
	psPlanoDestimo = CStr(ServiceVar("psPlanoDestino") )



	If Not ( psCarencia = "" ) Then
		psCarencia = Replace( Replace( psCarencia, "&lt", ">" ), "&gt", "<")
	End If


   Dim BSBEN002Dll As Object

   Set BSBEN002Dll = CreateBennerObject("BSBEN002.Modulo")

	Dim SQL As Object

	Set SQL = NewQuery


    SQL.Add("SELECT PLANO , MODULO  		                    	 ")
	SQL.Add("  FROM SAM_CONTRATO_MOD         ")
	SQL.Add(" WHERE HANDLE = :HANDLE   			 ")

	SQL.ParamByName("HANDLE").AsInteger = pHContratoMod



	SQL.Active = True

	Dim viPlano As Integer
	Dim viModDestino As Integer

	viPlano = SQL.FieldByName("PLANO").AsInteger
	viModDestino  = SQL.FieldByName("MODULO").AsInteger


	SQL.Active = False
	SQL.Clear
	SQL.Add("SELECT DESCRICAO FROM SAM_PLANO WHERE HANDLE = :HANDLE")
	SQL.ParamByName("HANDLE").AsInteger = viPlano
	SQL.Active = True



	ServiceVar("psPlanoDestino") = SQL.FieldByName("DESCRICAO").AsString


   psCarenciaDestino = BSBEN002Dll.AtualizaCarencia(CurrentSystem, _
											pHContratoMod, _
											viPlano, _
										    CDate(psDataAdesao), _
											vlHModBeneficiario, _
											psCarencia)


    ServiceVar("psCarenciaDestino")   =  psCarenciaDestino

	Set SQL = Nothing
	Set BSBEN002Dll = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = psMensagem

End Sub
