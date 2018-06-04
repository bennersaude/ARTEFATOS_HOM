'HASH: 9C8C5AF7450558CC07CCB5991B757064
Option Explicit
'#Uses "*bsShowMessage"
'#Uses "*CredenciamentoDePrestadores"

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim vMsg As String
    Dim vCurrentQuery As BPesquisa

    Set vCurrentQuery = NewQuery
    vCurrentQuery.Add("SELECT ISS, CATEGORIA FROM SAM_PRESTADOR WHERE HANDLE = :PRESTADOR")

  		vCurrentQuery.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")


  	vCurrentQuery.Active = True

	vMsg = ValidarPermissoesBotaoIniciarCredenciamento(vCurrentQuery)

	If (vMsg <> "") Then
		CanContinue = False
		bsshowmessage( vMsg, "E")
	End If

	If WebMode Then
	    TIPOCREDENCIAMENTO.WebLocalWhere = "A.HANDLE " + _
	                                       "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_CATEGORIA WHERE TIPOPROCESSO = A.HANDLE AND CATEGORIA = " + vCurrentQuery.FieldByName("CATEGORIA").AsString + ") AND " + _
	                                       "A.HANDLE " + _
	                                       "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_ISS WHERE TIPOPROCESSO = A.HANDLE AND ISS = " + vCurrentQuery.FieldByName("ISS").AsString + ")"
	Else
		TIPOCREDENCIAMENTO.LocalWhere = "HANDLE " + _
	                                    "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_CATEGORIA WHERE TIPOPROCESSO = SAM_TIPOPROCESSOCREDENCTO.HANDLE AND CATEGORIA = " + vCurrentQuery.FieldByName("CATEGORIA").AsString + ") AND " + _
	                                    "HANDLE " + _
	                                    "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_ISS WHERE TIPOPROCESSO = SAM_TIPOPROCESSOCREDENCTO.HANDLE AND ISS = " + vCurrentQuery.FieldByName("ISS").AsString + ")"
	End If

	Set vCurrentQuery = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
    If CurrentQuery.FieldByName("DATACREDENCIAMENTO").IsNull Then

	   Dim TvCredenciarPrestadorBLL As CSBusinessComponent
	   Dim exigeDataCredenciamento As Boolean

	   Set TvCredenciarPrestadorBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.TvCredenciarPrestadorBLL, Benner.Saude.Prestadores.Business")
       TvCredenciarPrestadorBLL.AddParameter(pdtInteger, CInt(CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger))
       exigeDataCredenciamento = TvCredenciarPrestadorBLL.Execute("HabilitarDataCredenciamento")
       Set TvCredenciarPrestadorBLL = Nothing

       If (exigeDataCredenciamento) Then
         CanContinue = False
         bsShowMessage("Este tipo de processo exige a data de credenciamento!", "E")
         Exit Sub
       End If
    End If
    If WebMode Then
	    Dim vMsg As String


		Dim vTipoProcesso As Long
		Dim vInseriuProcesso As Long
		vInseriuProcesso = -1

		On Error GoTo Except
			vMsg = ""
		    vInseriuProcesso = InserirProcessoCredenciamentoInicial(CurrentQuery.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger, _
			  			CurrentQuery.FieldByName("DATACREDENCIAMENTO").AsDateTime, CurrentQuery.FieldByName("NOVAFILIAL").AsInteger, CurrentQuery.FieldByName("INCLUIRFASES").AsString)

		    GoTo Fim
		Except:
			vInseriuProcesso = -2
			vMsg = Err.Description
		Fim:
		On Error GoTo 0

		Select Case vInseriuProcesso
			Case 0
				bsShowMessage("Inclusão de Credenciamento Concluída!", "I")
				If VisibleMode Then
					RefreshNodesWithTable("SAM_PRESTADOR")
				End If
			Case -1
				CanContinue = False
				bsShowMessage("Processo cancelado!", "I")

			Case -2
				CanContinue = False
				vMsg = "Falha ao inserir processo de credenciamento: " + Chr(13) + vMsg
				bsShowMessage(vMsg, "E")
		End Select
	End If
End Sub

Public Sub TABLE_NewRecord()
    CurrentQuery.FieldByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
End Sub

Public Sub TIPOCREDENCIAMENTO_OnChange()
	Dim TvCredenciarPrestadorBLL As CSBusinessComponent

	Set TvCredenciarPrestadorBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.TvCredenciarPrestadorBLL, Benner.Saude.Prestadores.Business")
    TvCredenciarPrestadorBLL.AddParameter(pdtInteger, CInt(CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger))
    If (Not TvCredenciarPrestadorBLL.Execute("HabilitarDataCredenciamento")) Then
        DATACREDENCIAMENTO.ReadOnly = True
        CurrentQuery.FieldByName("DATACREDENCIAMENTO").Clear
    Else
        DATACREDENCIAMENTO.ReadOnly = False
    End If

    If (Not TvCredenciarPrestadorBLL.Execute("HabilitarNovaFilial")) Then
      	NOVAFILIAL.ReadOnly = True
        CurrentQuery.FieldByName("NOVAFILIAL").Clear
    Else
        NOVAFILIAL.ReadOnly = False
    End If

    Set TvCredenciarPrestadorBLL = Nothing
End Sub
