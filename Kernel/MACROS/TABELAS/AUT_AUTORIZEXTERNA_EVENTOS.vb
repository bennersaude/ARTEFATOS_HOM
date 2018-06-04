'HASH: 49AEA3330D0EC22BD8169A41E1345B49
Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim BSAte009 As Object
Dim viGrau   As Long
Dim SP       As Object
Dim SQL      As Object

   Set SQL = NewQuery

   SQL.Add("SELECT RECEBEDOR,          ")
   SQL.Add("       LOCALEXECUCAO,      ")
   SQL.Add("       BENEFICIARIO,       ")
   SQL.Add("       TIPOAUTORIZACAO     ")
   SQL.Add("  FROM AUT_AUTORIZEXTERNA  ")
   SQL.Add(" WHERE HANDLE = :HANDLE")
   SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("AUTORIZEXTERNA").AsInteger
   SQL.Active = True


    Set SP = NewStoredProc
	SP.AutoMode = True
	SP.Name = "BSAUT_BUSCAPACOTE"

	SP.AddParam("p_Recebedor",ptInput)
	SP.AddParam("p_LocalExecucao",ptInput)
	SP.AddParam("p_Evento",ptInput)
    SP.AddParam("p_Beneficiario",ptInput)
    SP.AddParam("p_DataAtendimento",ptInput)
    SP.AddParam("p_Grau",ptInput)
    SP.AddParam("p_HandleTipoAutoriz",ptInput)
    SP.AddParam("p_Chave",ptInput)
    SP.AddParam("p_Usuario",ptInput)
    SP.AddParam("p_QtdPacote",ptOutput)


    SP.ParamByName("p_Recebedor").AsInteger          = SQL.FieldByName("RECEBEDOR").AsInteger
    SP.ParamByName("p_LocalExecucao").AsInteger      = SQL.FieldByName("LOCALEXECUCAO").AsInteger

    SP.ParamByName("p_Evento").AsInteger             = CurrentQuery.FieldByName("EVENTO").AsInteger

    SP.ParamByName("p_Beneficiario").AsInteger       = SQL.FieldByName("BENEFICIARIO").AsInteger

    SP.ParamByName("p_DataAtendimento").AsDateTime   = ServerDate
    SP.ParamByName("p_Grau").AsInteger               = CurrentQuery.FieldByName("GRAU").AsInteger

    SP.ParamByName("p_HandleTipoAutoriz").AsInteger  = SQL.FieldByName("TIPOAUTORIZACAO").AsInteger

    SP.ParamByName("p_Chave").AsInteger              = CurrentQuery.FieldByName("HANDLE").AsInteger
    SP.ParamByName("p_Usuario").AsInteger            = CurrentUser
	SP.ExecProc

    viGrau = SP.ParamByName("p_QtdPacote").AsInteger
	'Set BSAte009 = CreateBennerObject("BSAte009.Rotinas")
	'viGrau = BSAte009.ChecarPacote(CurrentSystem,CurrentQuery.FieldByName("AUTORIZEXTERNA").AsInteger,CurrentQuery.FieldByName("EVENTO").AsInteger,0)
	If viGrau > 1 Then
		CancelDescription = "Existe mais de um item de custo do tipo 'Pacote'. Por favor, selecione um item de custo do tipo 'Pacote'."
		CanContinue = False
	End If
	'Set BSAte009 = Nothing
    SQL.Active = False
    SP.AutoMode = False


End Sub
