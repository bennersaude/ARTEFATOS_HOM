'HASH: B387B100F725643DCF75469F25979739
Option Explicit

'#Uses "*bsShowMessage"

Dim vAlteracao As Boolean

Public Sub BOTAOCANCELAR_OnClick()
	Dim QDocumento As Object
	Dim QDocumentoFatura As Object
	Dim QUpdate As Object

	Set QDocumento = NewQuery
	Set QDocumentoFatura = NewQuery
	Set QUpdate = NewQuery

	If (CurrentQuery.State = 2 Or CurrentQuery.State = 3) Then
		bsShowMessage("A rotina não pode estar em Edição/Inserção !", "E")
		GoTo fim1:
	End If

	If CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
		bsShowMessage("Esta rotina ainda não foi processada !", "E")
		GoTo fim1:
	End If

	QDocumento.Clear
	QDocumento.Add("SELECT BAIXADATA,ULTIMAROTINAARQUIVODOC, CANCDATA FROM SFN_DOCUMENTO WHERE ROTINADOC = :PHANDLE AND (BAIXADATA IS NOT NULL OR ULTIMAROTINAARQUIVODOC IS NOT NULL OR CANCDATA IS NOT NULL)")
	QDocumento.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
	QDocumento.Active = True

	If Not QDocumento.FieldByName("CANCDATA").IsNull Then
		bsShowMessage("Já existe documento cancelado, impossível cancelar o parcelamento !", "E")
		GoTo fim1:
	End If

	If Not QDocumento.FieldByName("BAIXADATA").IsNull Then
		bsShowMessage("Já existe documento baixado, impossível cancelar o parcelamento !", "E")
		GoTo fim1:
	End If

	If Not QDocumento.FieldByName("ULTIMAROTINAARQUIVODOC").IsNull Then
		bsShowMessage("Documento já foi enviado em uma rotina arquivo, impossível cancelar o parcelamento !", "E")
		GoTo fim1:
	End If

	StartTransaction

	On Error GoTo FIM1:

	QDocumento.Clear
	QDocumento.Add("SELECT DISTINCT DOCUMENTOORIGINAL FROM SFN_DOCUMENTO WHERE ROTINADOC = :PHANDLE")
	QDocumento.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
	QDocumento.Active = True

	QDocumentoFatura.Clear
	QDocumentoFatura.Add("DELETE FROM SFN_DOCUMENTO_FATURA WHERE EXISTS (SELECT HANDLE FROM SFN_DOCUMENTO D WHERE D.HANDLE = SFN_DOCUMENTO_FATURA.DOCUMENTO AND D.ROTINADOC =:PROTINADOC)")
	QDocumentoFatura.ParamByName("PROTINADOC").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
	QDocumentoFatura.ExecSQL

	QDocumentoFatura.Clear
	QDocumentoFatura.Add("DELETE FROM SFN_DOCUMENTO WHERE ROTINADOC =:PROTINADOC")
	QDocumentoFatura.ParamByName("PROTINADOC").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
	QDocumentoFatura.ExecSQL


    QUpdate.Add("UPDATE SFN_DOCUMENTO SET CANCDATA =:CANCDATA, CANCMOTIVO =:CANCMOTIVO, ROTINADOC =:ROTINADOC WHERE HANDLE =:PHANDLE")

	While Not QDocumento.EOF
		    QUpdate.ParamByName("CANCDATA").DataType = ftDateTime
			QUpdate.ParamByName("CANCDATA").Clear
		    QUpdate.ParamByName("CANCMOTIVO").DataType = ftString
			QUpdate.ParamByName("CANCMOTIVO").Clear
			QUpdate.ParamByName("PHANDLE").AsInteger = QDocumento.FieldByName("DOCUMENTOORIGINAL").AsInteger
			QUpdate.ParamByName("ROTINADOC").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
			QUpdate.ExecSQL

			QDocumento.Next
	Wend


	vAlteracao = True
	CurrentQuery.Edit
	CurrentQuery.FieldByName("USUARIOCANCELAMENTO").AsInteger = CurrentUser
	CurrentQuery.FieldByName("DATAHORACANCELAMENTO").AsDateTime = ServerNow
	CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").Clear
	CurrentQuery.FieldByName("DATAHORAPROCESSAMENTO").Clear
	CurrentQuery.Post
	vAlteracao = False

	Commit


	bsShowMessage("Cancelamento efetuado !", "I")

	RefreshNodesWithTable("SFN_ROTINADOC")

	FIM1:
		If InTransaction Then Rollback

		If Error <> "" Then bsShowMessage(Str(Error), "E")

		Set QDocumentoFatura = Nothing
		Set QDocumento = Nothing
		Set QUpdate = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
	Dim QDocumento As Object
	Dim QDocumentoFatura As Object
	Dim QDocumentoFaturaCount As Object
	Dim QDocumentoFaturaSoma As Object
	Dim QInsere As Object
	Dim QInsereDocFat As Object
	Dim QUpdate As Object
	Dim vHandleDocumento As Long

	Set QDocumento = NewQuery
	Set QDocumentoFatura = NewQuery
	Set QDocumentoFaturaCount = NewQuery
	Set QDocumentoFaturaSoma = NewQuery
	Set QInsere = NewQuery
	Set QInsereDocFat = NewQuery
	Set QUpdate = NewQuery

	Dim vValorTotal As Double
	Dim vTotalSaldo As Double
	Dim vValorDocumento As Double
	Dim vValorParcela As Double
	Dim vPrimeiroHandleDocumento As Long
	Dim vNatureza As String
	Dim vNossoNumero As String

	Dim	I, J As Integer
	Dim vNumPar As Integer

	Dim Interface As Object

	Set Interface = CreateBennerObject("Financeiro.Documento")

	If (CurrentQuery.State = 2 Or CurrentQuery.State = 3) Then
		bsShowMessage("A rotina não pode estar em Edição/Inserção !", "E")
		GoTo fim:
	End If

	If Not CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
		bsShowMessage("Esta rotina já foi processada !", "I")
		GoTo fim:
	End If

	QDocumento.Clear
	QDocumento.Add("SELECT NUMERO,BAIXADATA,ULTIMAROTINAARQUIVODOC FROM SFN_DOCUMENTO WHERE ROTINADOC = :PHANDLE AND (BAIXADATA IS NOT NULL OR ULTIMAROTINAARQUIVODOC IS NOT NULL)")
	QDocumento.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
	QDocumento.Active = True

	If Not QDocumento.FieldByName("BAIXADATA").IsNull Then
		bsShowMessage("Já existe documento baixado, impossível efetuar o parcelamento !", "E")
		GoTo fim:
	End If

	If Not QDocumento.FieldByName("ULTIMAROTINAARQUIVODOC").IsNull Then
		bsShowMessage("Documento número "+QDocumento.FieldByName("NUMERO").AsString+" já foi enviado em uma rotina arquivo, impossível efetuar o parcelamento !", "I")
		GoTo fim:
	End If


	vNumPar = CurrentQuery.FieldByName("NUMERODEPARCELA").AsInteger

	If vNumPar <= 1 Then
		bsShowMessage("Informe um número maior que '1' para parcelamento !", "E")
		GoTo fim:
	End If

	QInsereDocFat.Add("INSERT INTO SFN_DOCUMENTO_FATURA")
	QInsereDocFat.Add("(HANDLE,DOCUMENTO,FATURA,SALDO,NATUREZA,VALORTOTAL)")
	QInsereDocFat.Add("VALUES")
	QInsereDocFat.Add("(:HANDLE,:DOCUMENTO,:FATURA,:SALDO,:NATUREZA,:VALORTOTAL)")


	QUpdate.Add("UPDATE SFN_DOCUMENTO SET CANCDATA =:CANCDATA, CANCMOTIVO =:CANCMOTIVO, ROTINADOC =:ROTINADOC WHERE HANDLE =:PHANDLE")

	QDocumento.Clear
	QDocumento.Add("SELECT * FROM SFN_DOCUMENTO WHERE ROTINADOC = :PHANDLE AND CANCDATA IS NULL AND DOCUMENTOORIGINAL IS NULL")
	QDocumento.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
	QDocumento.Active = True

	StartTransaction

	On Error GoTo FIM:

	vPrimeiroHandleDocumento = 0
	While Not QDocumento.EOF


			vValorParcela = QDocumento.FieldByName("VALOR").AsFloat / vNumPar

			QUpdate.ParamByName("CANCDATA").AsDateTime = ServerDate
			QUpdate.ParamByName("CANCMOTIVO").AsString = "Documento parcelado"
			QUpdate.ParamByName("PHANDLE").AsInteger = QDocumento.FieldByName("HANDLE").AsInteger
			QUpdate.ParamByName("ROTINADOC").DataType = ftInteger
			QUpdate.ParamByName("ROTINADOC").Clear
			QUpdate.ExecSQL

			QDocumentoFatura.Clear
			QDocumentoFatura.Add("SELECT * FROM SFN_DOCUMENTO_FATURA WHERE DOCUMENTO = :PDOCUMENTO")
			QDocumentoFatura.ParamByName("PDOCUMENTO").AsInteger = QDocumento.FieldByName("HANDLE").AsInteger
			QDocumentoFatura.Active = True

			QDocumentoFaturaCount.Clear
			QDocumentoFaturaCount.Add("SELECT COUNT(*) NUMREG FROM SFN_DOCUMENTO_FATURA WHERE DOCUMENTO = :PDOCUMENTO")
			QDocumentoFaturaCount.ParamByName("PDOCUMENTO").AsInteger = QDocumento.FieldByName("HANDLE").AsInteger
			QDocumentoFaturaCount.Active = True


    	vValorTotal = 0
		For I = 1 To vNumPar

		    vNossoNumero = GeraNossoNumero(QDocumento.FieldByName("TIPODOCUMENTO").AsInteger)
		    vHandleDocumento =	Interface.Criar(CurrentSystem, _
			                QDocumento.FieldByName("CONTAFINANCEIRA").AsInteger, _
			                QDocumento.FieldByName("TIPODOCUMENTO").AsInteger, _
			                QDocumento.FieldByName("DATAEMISSAO").AsDateTime, _
			                DateAdd("m",I-1,QDocumento.FieldByName("DATAVENCIMENTO").AsDateTime), _
			                DateAdd("m",I-1,QDocumento.FieldByName("COMPETENCIA").AsDateTime), _
			                "0", _
			                vNossoNumero, _
			                QDocumento.FieldByName("TESOURARIA").AsInteger, _
			                QDocumento.FieldByName("ROTINADOC").AsInteger, _
			                QDocumento.FieldByName("REGRAFINANCEIRA").AsInteger, _
			                QDocumento.FieldByName("CONTRATO").AsInteger, _
   			                QDocumento.FieldByName("FAMILIA").AsInteger, _
   			                QDocumento.FieldByName("FOLHAPAGAMENTO").AsInteger, _
   			                QDocumento.FieldByName("CODIGOFOLHA").AsInteger)



		    If vPrimeiroHandleDocumento = 0 Then
		        vPrimeiroHandleDocumento = vHandleDocumento
		    End If

			If I = vNumPar Then
			   AtualizaDocumento vHandleDocumento, Round(QDocumento.FieldByName("VALOR").AsFloat - vValorTotal,2), QDocumento.FieldByName("HANDLE").AsInteger, QDocumento.FieldByName("NATUREZA").AsString
    			vValorDocumento = Round(QDocumento.FieldByName("VALOR").AsFloat - vValorTotal,2)
			Else
			   AtualizaDocumento vHandleDocumento,Round(vValorParcela ,2), QDocumento.FieldByName("HANDLE").AsInteger, QDocumento.FieldByName("NATUREZA").AsString
    			vValorDocumento = Round(vValorParcela ,2)
			End If



		    vValorTotal = vValorTotal + Round(vValorParcela,2)

    		QDocumentoFatura.First
    		vTotalSaldo = 0
			For J = 1 To QDocumentoFaturaCount.FieldByName("NUMREG").AsInteger

				QInsereDocFat.ParamByName("HANDLE").AsInteger      = NewHandle("SFN_DOCUMENTO_FATURA")
				QInsereDocFat.ParamByName("DOCUMENTO").AsInteger   = vHandleDocumento
				QInsereDocFat.ParamByName("FATURA").AsInteger      = QDocumentoFatura.FieldByName("FATURA").AsInteger

		        If I = vNumPar Then
					QDocumentoFaturaSoma.Clear
					QDocumentoFaturaSoma.Add("SELECT SUM(SALDO) TOTAL FROM SFN_DOCUMENTO_FATURA WHERE FATURA = :PFATURA AND DOCUMENTO >=:PPRIMEIROHANDLE")
					QDocumentoFaturaSoma.ParamByName("PFATURA").AsInteger    = QDocumentoFatura.FieldByName("FATURA").AsInteger
					QDocumentoFaturaSoma.ParamByName("PPRIMEIROHANDLE").AsInteger =  vPrimeiroHandleDocumento
					QDocumentoFaturaSoma.Active = True 'SE FOR O ÚLTIMO DOCUMENTO VAI LANCANDO A DIFERENCA PARA CADA FATURA

    	    		QInsereDocFat.ParamByName("SALDO").AsFloat = Round(QDocumentoFatura.FieldByName("SALDO").AsFloat - QDocumentoFaturaSoma.FieldByName("TOTAL").AsFloat,2)


				ElseIf J = QDocumentoFaturaCount.FieldByName("NUMREG").AsInteger Then 'SE FOR O ULTIMO LANCAMENTO DO DOCUMENTO PARA A ULTIMA FATURA TEM QUE LANCAR A DIFERENCA
					If vTotalSaldo < 0 And QDocumento.FieldByName("NATUREZA").AsString = "D"  Then
			    		QInsereDocFat.ParamByName("SALDO").AsFloat = Abs(Round(vValorDocumento + vTotalSaldo,2))
			    	ElseIf vTotalSaldo > 0 And QDocumento.FieldByName("NATUREZA").AsString = "D"  Then
			    		QInsereDocFat.ParamByName("SALDO").AsFloat = Abs(Round(vValorDocumento + vTotalSaldo,2))
			    	ElseIf vTotalSaldo < 0 And QDocumento.FieldByName("NATUREZA").AsString = "C"  Then
			    		QInsereDocFat.ParamByName("SALDO").AsFloat = Abs(Round(vValorDocumento - vTotalSaldo,2))
			    	Else
			    		QInsereDocFat.ParamByName("SALDO").AsFloat = Abs(Round(vValorDocumento - vTotalSaldo,2))
			    	End If
			    Else
					QInsereDocFat.ParamByName("SALDO").AsFloat = Round(vValorDocumento/QDocumento.FieldByName("VALOR").AsFloat * QDocumentoFatura.FieldByName("SALDO").AsFloat,2)
				End If

				QInsereDocFat.ParamByName("NATUREZA").AsString     = QDocumentoFatura.FieldByName("NATUREZA").AsString
				QInsereDocFat.ParamByName("VALORTOTAL").AsFloat    = QInsereDocFat.ParamByName("SALDO").AsFloat
				QInsereDocFat.ExecSQL

				If QDocumentoFatura.FieldByName("NATUREZA").AsString = "C" Then
					vTotalSaldo = vTotalSaldo + Round(vValorDocumento/QDocumento.FieldByName("VALOR").AsFloat * QDocumentoFatura.FieldByName("SALDO").AsFloat,2)
				Else
					vTotalSaldo = vTotalSaldo - Round(vValorDocumento/QDocumento.FieldByName("VALOR").AsFloat * QDocumentoFatura.FieldByName("SALDO").AsFloat,2)
				End If
                QDocumentoFatura.Next
			Next J


		Next I

		QDocumento.Next
	Wend

	vAlteracao = True
	CurrentQuery.Edit
 	CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").AsInteger = CurrentUser
 	CurrentQuery.FieldByName("DATAHORAPROCESSAMENTO").AsDateTime = ServerNow
 	CurrentQuery.Post
	vAlteracao = False
	Commit

	bsShowMessage("Parcelamento concluído com êxito !", "I")

	RefreshNodesWithTable("SFN_ROTINADOC")

	FIM:

	If InTransaction Then Rollback

	If Error <> "" Then bsShowMessage("Ocorreu o seguinte erro no parcelamento "+Str(Error), "E")

	Set QDocumento = Nothing
	Set QDocumentoFatura = Nothing
	Set QDocumentoFaturaCount = Nothing
	Set QDocumentoFaturaSoma = Nothing
	Set QInsere = Nothing
	Set QInsereDocFat = Nothing
	Set QUpdate = Nothing
	Set Interface = Nothing

End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If Not CurrentQuery.FieldByName("DATAHORAPROCESSAMENTO").IsNull And Not vAlteracao Then
		bsShowMessage("Esta rotina já foi processada. Alteração não permitida !", "I")
		CanContinue = False
	End If
End Sub

Function GeraNossoNumero(pTipoDocto As Integer) As String
	Dim vErroFx As Boolean
	Dim LocHandleGerador As Integer

	Dim qUpdFaixa As Object
	Set qUpdFaixa = NewQuery

	Dim qGeradorNumero As Object
	Set qGeradorNumero = NewQuery

    qGeradorNumero.Clear
    qGeradorNumero.Add("SELECT GNF.HANDLE")
    qGeradorNumero.Add("FROM SFN_TIPODOCUMENTO TD, SFN_GERADORNUMERO GN, SFN_GERADORNUMERO_FAIXA GNF")
    qGeradorNumero.Add("WHERE GN.HANDLE=TD.GERADORNUMERO AND GNF.GERADORNUMERO=GN.HANDLE")
    qGeradorNumero.Add("      AND TD.TABTIPO = 3 AND GNF.ATUAL < GNF.FINAL AND TD.HANDLE=" + Str(pTipoDocto))
    qGeradorNumero.Add("ORDER BY GNF.ORDEM")
    qGeradorNumero.Active = True

    qUpdFaixa.Add("SELECT * FROM SFN_GERADORNUMERO_FAIXA WHERE HANDLE=:HANDLE")
    qUpdFaixa.RequestLive = True


    LocHandleGerador = qGeradorNumero.FieldByName("HANDLE").AsInteger
    If LocHandleGerador > 0 Then
        qUpdFaixa.Active =False
        qUpdFaixa.ParamByName("HANDLE").AsInteger = LocHandleGerador
        qUpdFaixa.Active = True

          qUpdFaixa.Edit
          qUpdFaixa.FieldByName("ATUAL").AsInteger = qUpdFaixa.FieldByName("ATUAL").AsInteger + 1
          qUpdFaixa.Post

          GeraNossoNumero = qUpdFaixa.FieldByName("ATUAL").AsString
    Else
      GeraNossoNumero = "0"
    End If
End Function

Public Sub AtualizaDocumento (pHandleDocumento As Long, pValorDocumento As Double, pHandleDocumentoOriginal As Long, pNatureza As String)
	Dim Query As Object
	Set Query = NewQuery

	Query.Add("UPDATE SFN_DOCUMENTO SET VALOR = :PVALOR, DOCUMENTOORIGINAL =:PDOCUMENTOORIGINAL, NATUREZA = :PNATUREZA WHERE HANDLE = :PHANDLE")
	Query.ParamByName("PDOCUMENTOORIGINAL").AsInteger = pHandleDocumentoOriginal
	Query.ParamByName("PVALOR").AsFloat = pValorDocumento
    Query.ParamByName("PNATUREZA").AsString = pNatureza
	Query.ParamByName("PHANDLE").AsInteger = pHandleDocumento
	Query.ExecSQL

	Set Query = Nothing
End Sub
