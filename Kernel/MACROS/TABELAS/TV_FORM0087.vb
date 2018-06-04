'HASH: C84FC65EBF245F77ACAAC6CDC2C3ED20
'#Uses "*bsShowMessage

Dim HandleContrato As String
Dim HandleTransferencia As Long

Public Sub DATAADESAO_OnChange()

If CurrentQuery.FieldByName("DATAADESAO").IsNull Then
	CurrentQuery.FieldByName("DATACANCELAMENTO").AsString = ""
Else
    CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime = CurrentQuery.FieldByName("DATAADESAO").AsDateTime - 1
End If

End Sub

Public Sub MODULODESTINO_OnChange()

Dim sql As Object
Set sql = NewQuery

sql.Clear
sql.Add("SELECT PLANO FROM SAM_CONTRATO_MOD WHERE CONTRATO = :CONTRATO AND MODULO = :MODULO")
sql.ParamByName("CONTRATO").AsString = HandleContrato
sql.ParamByName("modulo").AsString = CurrentQuery.FieldByName("MODULODESTINO").AsString
sql.Active = True

If CurrentQuery.FieldByName("MODULODESTINO").IsNull Then
	CurrentQuery.FieldByName("PLANODESTINO").AsString = ""

Else
	CurrentQuery.FieldByName("PLANODESTINO").AsInteger = sql.FieldByName("PLANO").AsInteger
End If

Set sql = Nothing

End Sub

Public Sub TABLE_AfterPost()
  If SessionVar("HANDLEBENEFICIARIO") <> "" Then
 	Exit Sub
  End If

  Dim ContratoHandle As Integer

  ContratoHandle = CStr(HandleContrato)

  	Dim vsMensagemErro As String
    Dim viRetorno As Long

    Dim vcContainer As CSDContainer
    Set vcContainer = NewContainer

    vcContainer.AddFields("HANDLE:INTEGER;CONTRATO:INTEGER;" + _
                          "MODULOORIGEM:INTEGER;PLANOORIGEM:INTEGER;MOTIVOCANCELAMENTO:INTEGER;DATACANCELAMENTO:DATETIME;" + _
                          "MODULODESTINO:INTEGER;PLANODESTINO:INTEGER;DATAADESAO:DATETIME")

    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = HandleTransferencia
    vcContainer.Field("CONTRATO").AsInteger = ContratoHandle
    vcContainer.Field("MODULOORIGEM").AsInteger = CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
    vcContainer.Field("PLANOORIGEM").AsInteger = CurrentQuery.FieldByName("PLANOORIGEM").AsInteger
    vcContainer.Field("MOTIVOCANCELAMENTO").AsInteger = CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsInteger
    vcContainer.Field("DATACANCELAMENTO").AsDateTime = CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime
    vcContainer.Field("MODULODESTINO").AsInteger = CurrentQuery.FieldByName("MODULODESTINO").AsInteger
    vcContainer.Field("PLANODESTINO").AsInteger = CurrentQuery.FieldByName("PLANODESTINO").AsInteger
    vcContainer.Field("DATAADESAO").AsDateTime = CurrentQuery.FieldByName("DATAADESAO").AsDateTime

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "BSBEN002", _
                                    "TransferirModuloContrato", _
                                    "Transferência de Modúlos do Contrato", _
                                     HandleTransferencia, _
                                     "SAM_CONTRATO_MOD_TRANSF", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     True, _
                                     vsMensagemErro, _
                                     vcContainer, _
                                     False)

     If viRetorno = 0 Then
	     bsShowMessage("Processo enviado para execução no servidor!", "I")
     Else
         bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
     End If

End Sub
Public Sub TABLE_AfterScroll()
	Dim sql As Object
	Set sql = NewQuery
	Dim sql2 As Object
	Set sql2 = NewQuery
	Dim sbWhere As String

	SessionVar("HANDLEBENEFICIARIO") = ""



	sql2.Add("SELECT * FROM SAM_CONTRATO WHERE CONTRATO = :PCONTRATO")
	sql2.ParamByName("PCONTRATO").AsString = SessionVar("CONTRATO")
	sql2.Active = True

	HandleContrato = sql2.FieldByName("HANDLE").AsString


	If CurrentQuery.State = 3 Then


		CONTRATO.Text = "Contrato : " + SessionVar("contrato") + " - " + SessionVar("contratante")
		CurrentQuery.FieldByName("MODULOORIGEM").AsString = SessionVar("modulo")
		CurrentQuery.FieldByName("PLANOORIGEM").AsString = SessionVar("plano")

		Dim Qry As Object
		Set Qry = NewQuery
		Dim sql1 As Object
		Set sql1 = NewQuery
		Dim qOrigem As Object
		Set qOrigem = NewQuery
		Dim QInsereModTransf As Object
		Set QInsereModTransf = NewQuery

		qOrigem.Active = False
		qOrigem.Clear
		qOrigem.Add("SELECT * FROM SAM_CONTRATO_MOD SCM")
		qOrigem.Add("JOIN SAM_REGISTROMS MS ON SCM.REGISTROMS = MS.HANDLE")
		qOrigem.Add("WHERE SCM.MODULO = :MODULO        ")
		qOrigem.Add("      AND SCM.CONTRATO = :CONTRATO")
		qOrigem.Add("      AND SCM.PLANO = :PLANO      ")
		qOrigem.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
		qOrigem.ParamByName("CONTRATO").AsInteger = HandleContrato
		qOrigem.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANOORIGEM").AsInteger
		qOrigem.Active = True

		Qry.Add("SELECT HANDLE, OBRIGATORIO FROM SAM_CONTRATO_MOD WHERE CONTRATO = :PCONTRATO AND MODULO = :PMODULO")
		Qry.ParamByName("PMODULO").AsString = SessionVar("modulo")
		Qry.ParamByName("PCONTRATO").AsString = HandleContrato
		Qry.Active = True

		HandleTransferencia = NewHandle("SAM_CONTRATO_MOD_TRANSF")
		QInsereModTransf.Active = False
		QInsereModTransf.Clear
		QInsereModTransf.Add("INSERT INTO SAM_CONTRATO_MOD_TRANSF ")
		QInsereModTransf.Add("  (HANDLE, CONTRATOMOD, SITUACAO)   ")
		QInsereModTransf.Add("VALUES                              ")
		QInsereModTransf.Add("  (:HANDLE, :CONTRATOMOD, :SITUACAO)")
		QInsereModTransf.ParamByName("HANDLE").AsInteger = HandleTransferencia
		QInsereModTransf.ParamByName("CONTRATOMOD").AsInteger = Qry.FieldByName("HANDLE").AsInteger
		QInsereModTransf.ParamByName("SITUACAO").AsString = "1"
		QInsereModTransf.ExecSQL
		sql1.Add("SELECT TIPOMODULO, TIPOCOBERTURA FROM SAM_MODULO WHERE HANDLE = :MODULO")
		sql1.ParamByName("MODULO").AsString = SessionVar("modulo")
		sql1.Active = True

        sbWhere = "HANDLE IN (SELECT SM.HANDLE FROM SAM_CONTRATO_MOD SCM " + _
									   "JOIN SAM_PLANO SP ON (SP.HANDLE = SCM.PLANO) " + _
									   "JOIN SAM_MODULO SM ON (SM.HANDLE = SCM.MODULO) " + _
									   "JOIN SAM_REGISTROMS SR ON (SR.HANDLE = SCM.REGISTROMS) " + _
									   "WHERE SCM.CONTRATO = " + HandleContrato + _
									   " AND SCM.MODULO <> " + SessionVar("MODULO") + _
									   "AND SCM.OBRIGATORIO = '" + Qry.FieldByName("OBRIGATORIO").AsString + "'" + _
									   " AND SCM.DATACANCELAMENTO IS NULL " + _
									   "AND SM.TIPOMODULO = '" + sql1.FieldByName("TIPOMODULO").AsString +"'" + _
									   " AND (SM.TIPOCOBERTURA = " + sql1.FieldByName("TIPOCOBERTURA").AsString + " OR SM.TIPOCOBERTURA = 3) "

			If qOrigem.FieldByName("NOVAREGULAMENTACAO").AsString = "S" Then
				sbWhere = sbWhere + " AND SR.NOVAREGULAMENTACAO = 'S') "
			Else
				sbWhere = sbWhere + " AND SR.NOVAREGULAMENTACAO IN ('S','N'))"
			End If

			If VisibleMode Then
				MODULODESTINO.LocalWhere = sbWhere
			ElseIf WebMode Then
				MODULODESTINO.WebLocalWhere	= sbWhere
			End If

			Set Qry = Nothing
			Set sql1 = Nothing
			Set qOrigem = Nothing

		Else
		If CurrentQuery.State = 3 Then

			CONTRATO.Text = "Beneficiário: "+ SessionVar("HANDLEBENEFICIARIO") + " - " + SessionVar("nomeBeneficiario")
			CurrentQuery.FieldByName("MODULOORIGEM").AsString = SessionVar("HANDLEPLANOMODULO")
			CurrentQuery.FieldByName("PLANOORIGEM").AsString = SessionVar("HANDLEMODULO")
		End If


	End If
	Set sql2 = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If SessionVar("HANDLEBENEFICIARIO") = "" Then

	Dim qDestino As Object
	Set qDestino = NewQuery
	Dim qOrigem As Object
	Set qOrigem = NewQuery


	If WebMode Then
		CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime = CurrentQuery.FieldByName("DATAADESAO").AsDateTime - 1
	    CurrentQuery.FieldByName("PLANODESTINO").AsInteger = CurrentQuery.FieldByName("PLANOORIGEM").AsInteger
	End If

	qDestino.Active = False
	qDestino.Clear
	qDestino.Add("SELECT * FROM SAM_CONTRATO_MOD")
	qDestino.Add("WHERE MODULO = :MODULO        ")
	qDestino.Add("      AND CONTRATO = :CONTRATO")
	qDestino.Add("      AND PLANO = :PLANO      ")
	qDestino.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULODESTINO").AsInteger
	qDestino.ParamByName("CONTRATO").AsInteger = HandleContrato
	qDestino.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANODESTINO").AsInteger
	qDestino.Active = True

	qOrigem.Active = False
	qOrigem.Clear
	qOrigem.Add("SELECT * FROM SAM_CONTRATO_MOD")
	qOrigem.Add("WHERE MODULO = :MODULO        ")
	qOrigem.Add("      AND CONTRATO = :CONTRATO")
	qOrigem.Add("      AND PLANO = :PLANO      ")
	qOrigem.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULOORIGEM").AsInteger
	qOrigem.ParamByName("CONTRATO").AsInteger = HandleContrato
	qOrigem.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANOORIGEM").AsInteger
	qOrigem.Active = True

	If Not qOrigem.FieldByName("DATACANCELAMENTO").IsNull Then
		bsShowMessage("Não é possível fazer esta transferência. Este contrato está cancelado!","E")
	    CanContinue = False
	End If

	If CurrentQuery.FieldByName("DATAADESAO").AsDateTime < qDestino.FieldByName("DATAADESAO").AsDateTime Then
		bsShowMessage("A data de adesão informada deve ser posterior à adesão do módulo destino.", "E")
	    CanContinue = False
	End If

	If Not qDestino.FieldByName("DATACANCELAMENTO").IsNull Then
		bsShowMessage("Módulo Destino não pode estar Cancelado.", "E")
	    CanContinue = False
	End If

	If ((qOrigem.FieldByName("REGISTROMS").IsNull) And (Not qDestino.FieldByName("REGISTROMS").IsNull)) Or _
	   ((Not qOrigem.FieldByName("REGISTROMS").IsNull) And (qDestino.FieldByName("REGISTROMS").IsNull)) Then
		bsShowMessage("Não será possível transferir o módulo. As cofigurações Do Registro ministério saúde são incompativeis!", "E")
	    CanContinue = False
	End If

	Set qDestino = Nothing
	Set qOrigem = Nothing

  Else
	Dim sql As Object
	Set sql = NewQuery

	Dim mensagem As String
	mensagem = ""

	sql.Clear
	sql.Add("SELECT DATAADESAO FROM SAM_BENEFICIARIO_MOD WHERE BENEFICIARIO = :BENEFICIARIO AND MODULO = :MODULO AND DATACANCELAMENTO IS NULL")
	sql.ParamByName("BENEFICIARIO").AsString = SessionVar("HANDLEBENEFICIARIO")
	sql.ParamByName("MODULO").AsString = CurrentQuery.FieldByName("MODULOORIGEM").AsString
	sql.Active = True


	If sql.FieldByName("DATAADESAO").AsDateTime >= CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
		mensagem = "A adesão do módulo de destino deve ser maior que a do módulo de origem!"
		CanContinue = False
	End If

	Set sql = Nothing
	If CanContinue Then
	  	Set Obj = CreateBennerObject("BSBEN002.Modulo")
	  	mensagem = Obj.TransfereModulo(CurrentSystem, _
	                                 CurrentQuery.FieldByName("MODULODESTINO").AsInteger, _
	                                 CurrentQuery.FieldByName("PLANODESTINO").AsInteger, _
	                                 RecordHandleOfTable("SAM_BENEFICIARIO_MOD"), _
	                                 CurrentQuery.FieldByName("MODULOORIGEM").AsInteger, _
	                                 CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsInteger, _
	                                 CurrentQuery.FieldByName("CODIGOTABELAPRC").AsInteger, _
	                                 CurrentQuery.FieldByName("DATAADESAO").AsDateTime, _
	                                 "", _
	                                 CurrentQuery.FieldByName("NAOTEMCARENCIA").AsBoolean)

		If mensagem <> "Concluído com sucesso." Then
		 	CanContinue = False
		End If

	End If

	If CanContinue Then
		bsShowMessage(mensagem, "I")
	Else
		bsShowMessage(mensagem, "E")
	End If
  End If
End Sub
