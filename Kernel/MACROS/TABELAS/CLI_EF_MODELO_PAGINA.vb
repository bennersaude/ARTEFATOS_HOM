'HASH: 4A7250F6E963782F77300A9C2321D6A0
Option Explicit
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

    Dim DataAdesalDll As Object
    Set DataAdesalDll = NewStoredProc
    DataAdesalDll.AutoMode = True
	DataAdesalDll.Name = "BS_F750651A"
	DataAdesalDll.AddParam("P_CONTRATO",ptInput, ftInteger)
	DataAdesalDll.AddParam("P_FAMILIA",ptInput, ftInteger)
	DataAdesalDll.AddParam("P_BENEFICIARIO",ptInput, ftInteger)
	DataAdesalDll.AddParam("P_NOVADATAADESAO",ptInput, ftDate)
	DataAdesalDll.AddParam("P_DATAPRIMEIRAADESAONOVA",ptInput, ftDate)
	DataAdesalDll.AddParam("P_USUARIOPROC",ptInput, ftInteger)
	DataAdesalDll.AddParam("P_MENSAGEM",ptOutput, ftString)

	DataAdesalDll.ParamByName("P_CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
    DataAdesalDll.ParamByName("P_FAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
    DataAdesalDll.ParamByName("P_BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    DataAdesalDll.ParamByName("P_NOVADATAADESAO").AsDateTime = CurrentQuery.FieldByName("DATAADESAONOVA").AsDateTime
    DataAdesalDll.ParamByName("P_DATAPRIMEIRAADESAONOVA").AsDateTime = CurrentQuery.FieldByName("DATAPRIMEIRAADESAONOVA").AsDateTime
    DataAdesalDll.ParamByName("P_USUARIOPROC").AsInteger = CurrentUser

	DataAdesalDll.ExecProc

	If DataAdesalDll.ParamByName("P_MENSAGEM").AsString <> "" Then
		bsShowMessage(DataAdesalDll.ParamByName("P_MENSAGEM").AsString, "E")
		CanContinue = False
	Else
		bsShowMessage("Alteração de Adesão efetuada com sucesso!", "I")
	End If

	Set DataAdesalDll = Nothing

End Sub
