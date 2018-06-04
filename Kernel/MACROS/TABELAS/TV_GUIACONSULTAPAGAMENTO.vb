'HASH: 124EA5D113AC0B10463B4CA59CBE512C
'#uses "*CriaTabelaTemporariaSqlServer"
Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If WebMode Then
		Dim SPP As Object

		Set SPP = NewStoredProc

		SPP.AutoMode = True
		SPP.Name = "BSAUT_AUTORIZWEB" 'SMS 78470 - Danilo Resende
		SPP.AddParam("p_WebAutoriz",ptInput)        	'Int
		SPP.ParamByName("p_WebAutoriz").DataType   		= ftInteger
		SPP.AddParam("P_VERSAOTISS",ptInput)			'Int       'Gabriel
		SPP.ParamByName("P_VERSAOTISS").DataType   		= ftInteger
		SPP.AddParam("p_TipoOperacao",ptInput)      	'Int
		SPP.ParamByName("p_TipoOperacao").DataType 		= ftInteger
		SPP.AddParam("p_Autorizacao",ptInput)       	'Int
		SPP.ParamByName("p_Autorizacao").DataType  		= ftInteger
		SPP.AddParam("p_TipoTISS",ptInput)          	'Varchar(1)
		SPP.ParamByName("p_TipoTISS").DataType     		= ftString
		SPP.AddParam("p_Origem",ptInput)            	'Varchar(1)
		SPP.ParamByName("p_Origem").DataType       		= ftString
		SPP.AddParam("p_Usuario",ptInput)           	'Int
		SPP.ParamByName("p_Usuario").DataType           = ftInteger
		SPP.AddParam("p_NumeroAutorizacao", ptInput)    'Float
		SPP.ParamByName("p_NumeroAutorizacao").DataType = ftFloat
		SPP.AddParam("p_EhReembolso",ptInput)          	'Varchar(1)
		SPP.ParamByName("p_EhReembolso").DataType     		= ftString
		SPP.AddParam("p_Retorno",ptOutput)              'Varchar(100)
		SPP.ParamByName("p_Retorno").DataType           = ftString

		SPP.ParamByName("p_TipoOperacao").AsInteger = 110
		SPP.ParamByName("p_Autorizacao").AsInteger  = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger

		If WebVisionCode = "W_TV_GUIA_CONSULTA_PRE" Then
			SPP.ParamByName("p_TipoTISS").AsString      = "C"
		ElseIf WebVisionCode = "W_TV_GUIACONSULTAPAGAMENTO" Then 'SMS 94421 - Marcelo Barbosa - 11/03/2008
			SPP.ParamByName("p_TipoTISS").AsString      = "S"
		End If

		SPP.ParamByName("p_Origem").AsString        =  2   ' "W"  - Alterada de acordo com a procedure
		SPP.ParamByName("p_Usuario").AsInteger      = CurrentUser
		SPP.ParamByName("p_WebAutoriz").Value       = Null


        'Gabriel Inicio
		Dim sql As BPesquisa
		Set sql = NewQuery

		sql.Add("SELECT MAX(HANDLE) HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'")
		sql.Active = True

	    SPP.ParamByName("P_VERSAOTISS").AsInteger = sql.FieldByName("HANDLE").AsInteger
 		'Gabriel - Fim

		If Not InTransaction Then
			StartTransaction
		End If

		If InStr(SQLServer, "SQL") > 0 Then
			Dim SQLx As Object
			Set SQLx = NewQuery

			On Error GoTo TabelasTemporarias

			SQLx.Clear
			SQLx.Add("SELECT 1 FROM #TMP_ORIGEMCALCULO")
			SQLx.Active = True

			Set SQLx = Nothing

			GoTo Procedure

			TabelasTemporarias:
			CriaTabelaTemporariaSqlServer

			Set SQLx = Nothing
		End If

		Procedure:
		On Error GoTo Erro
		SPP.ExecProc

        If SPP.ParamByName("P_RETORNO").AsString <> "" Then
		  InfoDescription = SPP.ParamByName("P_RETORNO").AsString
		End If
		If InTransaction Then
			Commit
		End If

		Set SPP = Nothing
	End If

  Exit Sub
  Erro:
    InfoDescription = Err.Description
    CancelDescription = Err.Description
    If InTransaction Then
      Rollback
    End If
End Sub


 
