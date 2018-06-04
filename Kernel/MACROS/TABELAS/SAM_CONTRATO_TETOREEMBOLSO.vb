'HASH: C58B193868D60BCC513F0C24C2537C05
'#Uses "*bsShowMessage"

Option Explicit

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
    Dim componente As CSBusinessComponent
    Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.Contrato.SamContratoTetoReembolsoBLL, Benner.Saude.Beneficiarios.Business")

    Dim vMensagem As String

    componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("CONTRATO").AsInteger)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("PLANO").AsInteger)
    componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)
    componente.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
    If (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    	componente.AddParameter(pdtString, CurrentQuery.FieldByName("DATAFINAL").AsString)
    Else
    	componente.AddParameter(pdtString, CurrentQuery.FieldByName("DATAFINAL").AsString)

	End If
	
    vMensagem = componente.Execute("VerificaDuplicidadeTetoReembolso")

    If vMensagem <> "" Then
        BsShowMessage(vMensagem,"E")
        CanContinue = False
    End If

    Set componente = Nothing
End Sub
