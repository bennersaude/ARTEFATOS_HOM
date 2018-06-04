'HASH: E9501EBAD3B005076A1963E3AA7EFDB1
'#Uses "*bsShowMessage"

Option Explicit

Dim vDias As Long
Dim vMeses As Long
Dim vAno As Long


Public Sub BOTAOCALCULAR_OnClick()
	Dim Obj As Object
	Set Obj = CreateBennerObject("Financeiro.Geral")

	Dim vJuros As Double
	Dim vMulta As Double
	Dim vCorrecao As Double
	Dim vDesconto As Double



	Obj.financeira(CurrentSystem, _
	               CurrentQuery.FieldByName("REGRAFINANCEIRA").AsInteger, _
	               CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime, _
	               CurrentQuery.FieldByName("DATABASECALCULO").AsDateTime, _
	               CurrentQuery.FieldByName("VALOR").AsCurrency, _
	               "", _
	               0, _
	               vJuros, _
	               vMulta, _
	               vCorrecao, _
	               vDesconto)


	CurrentQuery.FieldByName("JUROS").AsCurrency = vJuros
	CurrentQuery.FieldByName("MULTA").AsCurrency = vMulta
	CurrentQuery.FieldByName("CORRECAO").AsCurrency = vCorrecao
	CurrentQuery.FieldByName("DESCONTO").AsCurrency = vDesconto

	CurrentQuery.FieldByName("TOTAL").AsCurrency  =  CurrentQuery.FieldByName("VALOR").AsCurrency + vJuros + vMulta + vCorrecao - vDesconto

    Obj.DiferencaData(CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime, _
                      CurrentQuery.FieldByName("DATABASECALCULO").AsDateTime, _
                      vDias, _
                      vMeses, _
                      vAno)

	ROTULOCALCULO.Text = "Cálculo realizado: " & FormatDateTime2("DD/MM/YYYY hh:nn:ss",ServerNow) & vbNewLine & _
		"Período calculado: " & Str(vAno) & IIf(vAno<=1," Ano,  "," Anos,  ")& Str(vMeses) & IIf(vMeses<=1," Mes e"," Meses e ") & Str(vDias) & IIf(vDias<=1," Dia."," Dias.")


	Set Obj = Nothing
End Sub

Public Sub TABLE_AfterInsert()
	If UserVar("HANDLE_REGRAFINANCEIRA") <> "" Then
		CurrentQuery.FieldByName("REGRAFINANCEIRA").AsInteger = CInt(UserVar("HANDLE_REGRAFINANCEIRA"))
        UserVar("HANDLE_REGRAFINANCEIRA") = ""
	End If
	CurrentQuery.FieldByName("DATABASECALCULO").AsDateTime = ServerDate

End Sub

Public Sub TABLE_AfterPost()
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	BOTAOCALCULAR_OnClick
	If VisibleMode Then
		CanContinue = False
	Else
		bsShowMessage("Cálculo realizado: " & FormatDateTime2("DD/MM/YYYY hh:nn:ss",ServerNow) & vbNewLine & _
		"Total: " & Format(CurrentQuery.FieldByName("TOTAL").AsCurrency,"#,###0.00") & vbNewLine & "Período calculado: " & Str(vAno) & IIf(vAno<=1," Ano,  "," Anos,  ")& Str(vMeses) & IIf(vMeses<=1," Mes e"," Meses e ") & Str(vDias) & IIf(vDias<=1," Dia."," Dias.") ,"I")

	End If

End Sub
