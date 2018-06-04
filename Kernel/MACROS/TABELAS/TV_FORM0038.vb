'HASH: 4DC016564172DB98B47C4DCCCD6414EB
'#Uses "*bsShowMessage"
Dim giOperacaoEspecial As Long


Public Sub CalculaValorJurosMulta()

  Dim vfJuro As Double
  Dim vfMulta As Double
  Dim vfCorrecao As Double
  Dim vfDesconto As Double
  Dim FINANCEIRO As Object
  Set FINANCEIRO = CreateBennerObject("FINANCEIRO.Geral")
  FINANCEIRO.Financeira(CurrentSystem, _
	  				    CurrentQuery.FieldByName("REGRAFINANCEIRAFATURA").AsInteger, _
 					    CurrentQuery.FieldByName("DATAVENCIMENTOFATURA").AsDateTime, _
					    CurrentQuery.FieldByName("DATABAIXA").AsDateTime, _
					    CurrentQuery.FieldByName("VALORCALC").AsFloat, _
					    CurrentQuery.FieldByName("NATUREZAFATURA").AsString, _
					    0, _
					    vfJuro, _
					    vfMulta, _
					    vfCorrecao, _
					    vfDesconto)

  Set FINANCEIRO = Nothing

  CurrentQuery.FieldByName("JUROCALC").AsFloat = vfJuro
  CurrentQuery.FieldByName("MULTACALC").AsFloat = vfMulta
  CurrentQuery.FieldByName("CORRECAOCALC").AsFloat = vfCorrecao
  CurrentQuery.FieldByName("DESCONTOCALC").AsFloat = vfDesconto
  CurrentQuery.FieldByName("JUROINFORMADO").AsFloat = CurrentQuery.FieldByName("JUROCALC").AsFloat
  CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat = CurrentQuery.FieldByName("MULTACALC").AsFloat
  CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat = CurrentQuery.FieldByName("CORRECAOCALC").AsFloat
  CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat =	CurrentQuery.FieldByName("DESCONTOCALC").AsFloat

  CurrentQuery.FieldByName("TOTALCALC").AsFloat = CurrentQuery.FieldByName("VALORCALC").AsFloat + _
	  											  CurrentQuery.FieldByName("JUROCALC").AsFloat + _
	 											  CurrentQuery.FieldByName("MULTACALC").AsFloat + _
												  CurrentQuery.FieldByName("CORRECAOCALC").AsFloat - _
												  CurrentQuery.FieldByName("DESCONTOCALC").AsFloat
  CurrentQuery.FieldByName("TOTALINFORMADO").AsFloat = CurrentQuery.FieldByName("VALORINFORMADO").AsFloat + _
													   CurrentQuery.FieldByName("JUROINFORMADO").AsFloat + _
													   CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat + _
													   CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat - _
													   CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat
  CalculaValor

End Sub


'INICIO - SMS 131975 - GUSTAVO GALINA - 09/06/2010
Public Sub CalculaValor()                           ' Calcula o total a cada alteração de algum dos campos
	CurrentQuery.FieldByName("TOTALINFORMADO").AsFloat =	CurrentQuery.FieldByName("VALORINFORMADO").AsFloat + CurrentQuery.FieldByName("JUROINFORMADO").AsFloat     + _
															CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat + CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat - _
															CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat
End Sub

Public Sub CORRECAOINFORMADO_OnExit()

	CalculaValor

End Sub

Public Sub DATABAIXA_OnExit()

  CalculaValorJurosMulta

End Sub

Public Sub DESCONTOINFORMADO_OnExit()

	CalculaValor

End Sub

Public Sub JUROINFORMADO_OnExit()

	CalculaValor

End Sub


Public Sub MULTAINFORMADO_OnExit()

	CalculaValor

End Sub
'FIM - SMS 131975

Public Sub TABLE_AfterInsert()
	Dim qFatura As BPesquisa
	Dim viDias As Long
	Dim vfJuro As Double
	Dim vfMulta As Double
	Dim vfCorrecao As Double
	Dim vfDesconto As Double
	Dim FINANCEIRO As Object
	Set qFatura = NewQuery
	Set FINANCEIRO = CreateBennerObject("FINANCEIRO.Geral")

	qFatura.Clear
	qFatura.Add("SELECT HANDLE,			  ")
	qFatura.Add("		VALORRETENCAOIRRF,")
	qFatura.Add("		VALORRETENCAOISS, ")
	qFatura.Add("		NUMERO,			  ")
	qFatura.Add("		TIPOFATURAMENTO,  ")
	qFatura.Add("		DATAEMISSAO,	  ")
	qFatura.Add("		DATACONTABIL,	  ")
	qFatura.Add("		DATAVENCIMENTO,	  ")
	qFatura.Add("		VALOR,			  ")
	qFatura.Add("		SALDO,			  ")
	qFatura.Add("		NATUREZA,		  ")
	qFatura.Add("		REGRAFINANCEIRA   ")
	qFatura.Add("  FROM SFN_FATURA		  ")
	qFatura.Add(" WHERE HANDLE = :HANDLE  ")
	qFatura.Add("	AND SALDO > 0		  ")
	qFatura.ParamByName("HANDLE").AsString = SessionVar("HFATURA")
	qFatura.Active = True

	CurrentQuery.FieldByName("NUMEROFATURA").AsInteger = qFatura.FieldByName("NUMERO").AsInteger
	CurrentQuery.FieldByName("TIPOFATURAMENTOFATURA").AsInteger = qFatura.FieldByName("TIPOFATURAMENTO").AsInteger
	CurrentQuery.FieldByName("DATAEMISSAOFATURA").AsDateTime = qFatura.FieldByName("DATAEMISSAO").AsDateTime
	CurrentQuery.FieldByName("DATACONTABIL").AsDateTime = qFatura.FieldByName("DATACONTABIL").AsDateTime ' SUGERIR O MESMO DA FATURA
	CurrentQuery.FieldByName("DATACONTABILFATURA").AsDateTime = qFatura.FieldByName("DATACONTABIL").AsDateTime
	CurrentQuery.FieldByName("DATAVENCIMENTOFATURA").AsDateTime = qFatura.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("VALORFATURA").AsFloat = qFatura.FieldByName("VALOR").AsFloat
	CurrentQuery.FieldByName("SALDOFATURA").AsFloat = qFatura.FieldByName("SALDO").AsFloat
	CurrentQuery.FieldByName("NATUREZAFATURA").AsString = qFatura.FieldByName("NATUREZA").AsString
	CurrentQuery.FieldByName("REGRAFINANCEIRAFATURA").AsInteger = qFatura.FieldByName("REGRAFINANCEIRA").AsInteger

	If qFatura.EOF Then
		bsShowMessage("Fatura não possui saldo a ser baixado!", "I")
	Else
		CurrentQuery.FieldByName("VALORCALC").AsFloat = qFatura.FieldByName("SALDO").AsFloat
		CurrentQuery.FieldByName("VALORINFORMADO").AsFloat = qFatura.FieldByName("SALDO").AsFloat

		viDias = DateDiff("y", qFatura.FieldByName("DATAVENCIMENTO").AsDateTime, CurrentQuery.FieldByName("DATABAIXA").AsDateTime)

		CurrentQuery.FieldByName("DIFERENCADATAVCTOBAIXA").AsString = CStr(viDias) + " dia(s)"

        CalculaValorJurosMulta
	End If
End Sub

Public Sub TABLE_AfterPost()
	Dim SFNBAIXA As Variant
	Dim viRetorno As Long
	Dim vsMensagem As String
	Set SFNBAIXA = CreateBennerObject("SFNBAIXA.Documento")

	viRetorno = SFNBAIXA.BaixarFatura(CurrentSystem, _
									  CLng(SessionVar("HFATURA")), _
									  CurrentQuery.FieldByName("DATABAIXA").AsDateTime, _
									  CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
									  CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
									  CurrentQuery.FieldByName("HISTORICO").AsString, _
									  CurrentQuery.FieldByName("VALORCALC").AsFloat, _
									  CurrentQuery.FieldByName("VALORINFORMADO").AsFloat, _
									  CurrentQuery.FieldByName("JUROINFORMADO").AsFloat, _
									  CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat, _
									  CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat, _
									  CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat, _
									  CurrentQuery.FieldByName("NATUREZAFATURA").AsString, _
									  0, _
									  giOperacaoEspecial, _
									  0, _
									  0, _
									  vsMensagem)


  If WebMode Then
  	If vsMensagem <> "" Then
  		bsShowMessage(vsMensagem, "I")
  	Else
		If viRetorno = 0 Then
		  	bsShowMessage("Baixa de fatura realizada com sucesso!", "I")
  		End If
  	End If

  ElseIf VisibleMode Then
	If viRetorno = 0 Then
	  	bsShowMessage("Baixa de fatura realizada com sucesso!", "I")
  	Else
		Err.Raise(vbsUserException, "", vsMensagem)
  	End If

  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SAMCONTAFINANCEIRA As Object
  Dim vsDocumentos As String
  Dim vsMensagemRetorno As String

  Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")

  If Not SAMCONTAFINANCEIRA.VerificaDocumentos(CurrentSystem,CLng(SessionVar("HFATURA")),vsDocumentos,vsMensagemRetorno) Then
    BsShowMessage(vsMensagemRetorno,"E")
    CanContinue = False
  End If

  Set SAMCONTAFINANCEIRA = Nothing

  Dim qTipoFatura As Object

  Set qTipoFatura = NewQuery
  qTipoFatura.Add("SELECT TIPOFATURAMENTO FROM SFN_FATURA WHERE HANDLE = " + SessionVar("HFATURA"))
  qTipoFatura.Active = True


  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + qTipoFatura.FieldByName("TIPOFATURAMENTO").AsString)
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
    bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
    CanContinue = False
  End If

  Set Sql = Nothing
  Set qTipoFatura = Nothing


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qRotina As BPesquisa
	Dim vbRotIRRF As Boolean
	Dim vbRotISS As Boolean
	Dim viOperacaoEspecial As Integer
	Set qRotina = NewQuery

	If (CurrentQuery.FieldByName("DATABAIXA").AsDateTime > CurrentQuery.FieldByName("DATACONTABIL").AsDateTime) Then
		bsShowMessage("Data contábil não pode ser menor que data para baixa!", "E")

		CanContinue = False

		Exit Sub
	End If

	qRotina.Clear
	qRotina.Add("SELECT FATURA ")
	qRotina.Add("  FROM SFN_ROTINAFIN_IRRF_CODDIRF_FAT")
	qRotina.Add(" WHERE FATURA = :NFATURA")
	qRotina.ParamByName("NFATURA").AsString = SessionVar("HFATURA")
	qRotina.Active = True

	vbRotIRRF = (Not qRotina.EOF)

	qRotina.Active = False
	qRotina.Clear
	qRotina.Add("SELECT FATURA")
	qRotina.Add("  FROM SFN_ROTINAFINISS_FATURA")
	qRotina.Add(" WHERE FATURA = :NFATURA")
	qRotina.ParamByName("NFATURA").AsString = SessionVar("HFATURA")
	qRotina.Active = True

	vbRotISS = (Not qRotina.EOF)

	If ((vbRotIRRF Or vbRotISS) And _
		(CurrentQuery.FieldByName("VALORCALC").AsFloat <> CurrentQuery.FieldByName("VALORINFORMADO").AsFloat)) Then
		bsShowMessage("Fatura com imposto que se encontra em rotina de recolhimento, " + Chr(13) + "não pode ser baixada parcialmente!", "E")

		CanContinue = False

		Exit Sub
	End If

	If (CurrentQuery.FieldByName("VALORINFORMADO").AsFloat > CurrentQuery.FieldByName("VALORCALC").AsFloat) Then
		bsShowMessage("O valor não pode ser maior que o calculado!", "E")

		CanContinue = False

		Exit Sub
	End If

	If (CurrentQuery.FieldByName("DATABAIXA").AsDateTime < CurrentQuery.FieldByName("DATAEMISSAOFATURA").AsDateTime) Then
		bsShowMessage("Data não pode ser inferior a data de emissão!", "E")

		CanContinue = False

		Exit Sub
	End If

	If ((CurrentQuery.FieldByName("TESOURARIA").AsInteger <= 0) And (Not (CurrentQuery.FieldByName("LUCROPERDA").AsString = "S"))) Then
		bsShowMessage("Tesouraria não informada!", "E")

		CanContinue = False

		Exit Sub
	End If

	If (CurrentQuery.FieldByName("LUCROPERDA").AsString = "S") Then
		giOperacaoEspecial = 131
	Else
		giOperacaoEspecial = 130
	End If
End Sub

Public Sub VALORINFORMADO_OnExit() 'SMS - 131975 - GUSTAVO GALINA - 09/06/2010
	Dim Razao As Double 'Calcula os campos para que fiquem proporcionais ao novo valor informado em relação ao calculado

	Razao = CurrentQuery.FieldByName("VALORINFORMADO").AsFloat / CurrentQuery.FieldByName("VALORCALC").AsFloat
    CurrentQuery.FieldByName("JUROINFORMADO").AsFloat = CurrentQuery.FieldByName("JUROCALC").AsFloat * Razao
    CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat = CurrentQuery.FieldByName("MULTACALC").AsFloat * Razao
    CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat = CurrentQuery.FieldByName("CORRECAOCALC").AsFloat * Razao
    CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat = CurrentQuery.FieldByName("DESCONTOCALC").AsFloat * Razao
    CurrentQuery.FieldByName("TOTALINFORMADO").AsFloat = CurrentQuery.FieldByName("TOTALCALC").AsFloat * Razao

End Sub
