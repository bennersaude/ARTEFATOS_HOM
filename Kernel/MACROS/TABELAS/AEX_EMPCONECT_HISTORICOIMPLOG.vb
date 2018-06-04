'HASH: C802F5B2F1615EB37E416E38662E707F
'----------------------------------------------------------------------------
'Macro AEX_LOGRETORNOARQUIVOINFO
' atualizada em 10/08/2007
'----------------------------------------------------------------------------
Public Sub BOTAOCANCELAR_OnClick()
		Dim SPP1 As Object
		Dim vRetorno

		Set SPP1 = NewStoredProc
		SPP1.Name = "BSAEX_CANCELALOGIMPORT"
		SPP1.AutoMode = True

        'Tratando os tipo dos parâmetros e se são de imput ou ouput
		SPP1.AddParam("P_HANDLEAEXLOGRETORNOARQINFO",ptInput)
		SPP1.ParamByName("P_HANDLEAEXLOGRETORNOARQINFO").DataType     = ftInteger

		SPP1.AddParam("P_USUARIO",ptInput)
		SPP1.ParamByName("P_USUARIO").DataType        = ftInteger

		SPP1.AddParam("P_RETORNO",ptOutput)
		SPP1.ParamByName("P_RETORNO").DataType = ftString


        'Passando os parâmetros para a SP
		SPP1.ParamByName("P_HANDLEAEXLOGRETORNOARQINFO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		SPP1.ParamByName("P_USUARIO").AsInteger = CurrentUser
		'Executando a SP
		SPP1.ExecProc

        vRetorno = SPP1.ParamByName("P_RETORNO").AsString
        
        MsgBox vRetorno
        RefreshNodesWithTable ("AEX_EMPCONECT_HISTORICOIMPLOG")
	End Sub

Public Sub TABLE_AfterPost()
   If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
     BOTAOCANCELAR.Enabled = True
   Else
     BOTAOCANCELAR.Enabled = False
   End If
End Sub

Public Sub TABLE_AfterScroll()
   If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
     BOTAOCANCELAR.Enabled = True
   Else
     BOTAOCANCELAR.Enabled = False
   End If
End Sub


