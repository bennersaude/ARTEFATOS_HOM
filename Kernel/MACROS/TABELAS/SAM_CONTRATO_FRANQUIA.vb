'HASH: 89FAE1CC7BF4EED38A46828E4985EE4F
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
    If WebMode Then
        PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
    ElseIf VisibleMode Then
		PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
        PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
    ElseIf VisibleMode Then
		PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = " AND FRANQUIA = " + CurrentQuery.FieldByName("FRANQUIA").AsString +" AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString   'Anderson sms 21638(PLANO)
  'Valeska -sms 25507 -Não poderá existir duas franquias iguais para o mesmo contrato,mesmo se forem de
  'planos diferentes,pois não tem como fazer a contagem para o beneficiário -Processamento de Contas

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_FRANQUIA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Set Interface = Nothing
    Exit Sub
  End If

End Sub

