'HASH: 123E5863CE0D9C0C786320FF3D8761E2
'Macro: SAM_TIPOCREDENCIAMENTO_FASE
'#Uses "*bsShowMessage"


Public Sub TABLE_AfterPost()

  Dim SQL As Object
  Dim vFrase As String

  If CurrentQuery.FieldByName("ULTIMAFASE").AsString = "S" Then

    Set SQL = NewQuery

    vFrase = "UPDATE SAM_TIPOCREDENCIAMENTO_FASE SET ULTIMAFASE = 'N' WHERE "
    vFrase = vFrase + "HANDLE <> :vHandle AND TIPOCREDENCIAMENTO = :vTipoCredencimanto"
    SQL.Clear
    SQL.Add(vFrase)
    SQL.ParamByName("vHandle").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ParamByName("vTipoCredencimanto").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger
    SQL.ExecSQL

    Set SQL = Nothing

  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		FASE.WebLocalWhere = "A.HANDLE IN (SELECT F.HANDLE 											" + _
										  "  FROM SAM_FASE F 										" + _
										  " WHERE NOT EXISTS (SELECT 1 								" + _
										  "                     FROM SAM_TIPOCREDENCIAMENTO_FASE TF " + _
										  "                    WHERE TF.FASE = F.HANDLE 			" + _
										  "                      AND TF.TIPOCREDENCIAMENTO = @CAMPO(TIPOCREDENCIAMENTO) ))"
	ElseIf VisibleMode Then
		FASE.LocalWhere = "HANDLE IN (SELECT F.HANDLE                                           " + _
                           	         "  FROM SAM_FASE F                                         " + _
                          			 " WHERE NOT EXISTS (SELECT 1                               " + _
                    	  			 "                     FROM SAM_TIPOCREDENCIAMENTO_FASE TF  " + _
                    				 "                    WHERE TF.FASE = F.HANDLE 				" + _
                    				 "                      AND TF.TIPOCREDENCIAMENTO = @TIPOCREDENCIAMENTO))"
    End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		FASE.WebLocalWhere = "A.HANDLE IN (SELECT F.HANDLE 											" + _
										  "  FROM SAM_FASE F 										" + _
										  " WHERE NOT EXISTS (SELECT 1 								" + _
										  "                     FROM SAM_TIPOCREDENCIAMENTO_FASE TF " + _
										  "                    WHERE TF.FASE = F.HANDLE 			" + _
										  "                      AND TF.TIPOCREDENCIAMENTO = @CAMPO(TIPOCREDENCIAMENTO) ))"
	ElseIf VisibleMode Then
		FASE.LocalWhere = "HANDLE IN (SELECT F.HANDLE                                           " + _
                           	         "  FROM SAM_FASE F                                         " + _
                          			 " WHERE NOT EXISTS (SELECT 1                               " + _
                    	  			 "                     FROM SAM_TIPOCREDENCIAMENTO_FASE TF  " + _
                    				 "                    WHERE TF.FASE = F.HANDLE 				" + _
                    				 "                      AND TF.TIPOCREDENCIAMENTO = @TIPOCREDENCIAMENTO))"
    End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	On Error GoTo Exception

	Dim componente As CSBusinessComponent
	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamTipoCredenciamentoFaseBLL, Benner.Saude.Prestadores.Business")
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("ORDEMRATIFICACAO").AsInteger)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("REGRAAPROVACAORATIFICACAO").AsInteger)
	componente.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("EXIGIRRATIFICACAO").AsBoolean)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("CONTROLARFASERATIFICACAO").Value)

	componente.Execute("VerificarControleRatificacao")
	Set componente = Nothing
	Exit Sub

	Exception:
    	Set componente = Nothing
    	bsShowMessage(Err.Description, "I")
    	CanContinue = False
    	Exit Sub
End Sub

Public Sub TABLE_NewRecord()

	CurrentQuery.FieldByName("CONTROLARFASERATIFICACAO").Value = 2

	Dim vQueryControlaRatificacao As BPesquisa
	Set vQueryControlaRatificacao = NewQuery

	vQueryControlaRatificacao.Add("SELECT CONTROLARRATIFICACAO           ")
	vQueryControlaRatificacao.Add("		  FROM SAM_TIPOPROCESSOCREDENCTO ")
	vQueryControlaRatificacao.Add("WHERE HANDLE = :PHANDLE               ")

	vQueryControlaRatificacao.ParamByName("PHANDLE").Value = CStr(RecordHandleOfTable("SAM_TIPOPROCESSOCREDENCTO"))
	vQueryControlaRatificacao.Active = True


	If (vQueryControlaRatificacao.FieldByName("CONTROLARRATIFICACAO").AsString = "S") Then

		CurrentQuery.FieldByName("CONTROLARFASERATIFICACAO").Value = 1

		Dim vQueryOrdem As BPesquisa
		Set vQueryOrdem = NewQuery

		vQueryOrdem.Add("SELECT COALESCE(MAX(ORDEMRATIFICACAO),0) PROXIMA ")
		vQueryOrdem.Add("		FROM SAM_TIPOCREDENCIAMENTO_FASE          ")
		vQueryOrdem.Add("WHERE TIPOCREDENCIAMENTO = :PROCESSO             ")
		vQueryOrdem.Add("AND CONTROLARFASERATIFICACAO = 1                 ")

		vQueryOrdem.ParamByName("PROCESSO").Value = CStr(RecordHandleOfTable("SAM_TIPOPROCESSOCREDENCTO"))
		vQueryOrdem.Active = True

		CurrentQuery.FieldByName("ORDEMRATIFICACAO").Value = vQueryOrdem.FieldByName("PROXIMA").AsInteger + 1

		Set vQueryOrdem = Nothing
	End If

	Set vQueryControlaRatificacao = Nothing
End Sub
