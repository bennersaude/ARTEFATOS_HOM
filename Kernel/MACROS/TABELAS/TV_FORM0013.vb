'HASH: 5D09B1D09EC4F8066EC1BF5C08DBC0B4
 
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_AfterScroll()
Dim HBENEFICIARIO        As Long

If VisibleMode Then
  IMPRIMIRCARTAO.Visible = False
End If

If SessionVar("HBENEFICIARIO") <> "" Then
  HBENEFICIARIO = CLng(SessionVar("HBENEFICIARIO"))
Else
  HBENEFICIARIO = RecordHandleOfTable("SAM_BENEFICIARIO")
End If

Dim SQL As Object

Set SQL = NewQuery
SQL.Add("SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
SQL.ParamByName("HANDLE").AsInteger = HBENEFICIARIO

SQL.Active = True

If VisibleMode Then
 	CARTAOMOTIVO.LocalWhere = "VISUALIZANAUNIDADE = 'S' And HANDLE IN (Select C.CARTAOMOTIVO  FROM SAM_CONTRATO_CARTAOMOTIVO C WHERE C.CONTRATO = " + SQL.FieldByName("CONTRATO").AsString + ")"
ElseIf WebMode Then
	CARTAOMOTIVO.WebLocalWhere = "A.VISUALIZANAUNIDADE = 'S' And A.HANDLE IN (Select C.CARTAOMOTIVO  FROM SAM_CONTRATO_CARTAOMOTIVO C WHERE C.CONTRATO = " + SQL.FieldByName("CONTRATO").AsString + ")"
End If

Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim HBENEFICIARIO        As Long
    Dim bs As CSBusinessComponent
    Dim resultado As String

    If SessionVar("HBENEFICIARIO") <> "" Then
    	HBENEFICIARIO = CLng(SessionVar("HBENEFICIARIO"))
    Else
        HBENEFICIARIO = RecordHandleOfTable("SAM_BENEFICIARIO")
    End If

    If HBENEFICIARIO = 0 Then
      bsShowMessage("Não foi selecionado um beneficiário!", "E")
      CanContinue = False
    Else
      Set bs = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.Cartao.EmissaoCartaoAvulso, Benner.Saude.Beneficiarios.Business") ' formato: [namespace.classe], [assembly]

      bs.ClearParameters
      bs.AddParameter(pdtInteger, HBENEFICIARIO)
      bs.AddParameter(pdtString, CurrentQuery.FieldByName("NAOCOBRAREMISSAO").AsString)
      bs.AddParameter(pdtString, "S")

      resultado = CStr(bs.Execute("ConsultaCobrancaEmissao"))

      If VisibleMode Then
        If resultado <> "" Then
          If bsShowMessage("Confirma não Faturar?", "Q") = vbNo Then
            CanContinue = False
            Exit Sub
          End If
        End If
      End If

      resultado = ""
      bs.ClearParameters
      bs.AddParameter(pdtInteger, HBENEFICIARIO)
      If WebMode Then
        bs.AddParameter(pdtString, "S")
      Else
        bs.AddParameter(pdtString, "N")
      bs.AddParameter(pdtString, SessionVar("MIGRACAOCONTRATO"))

      resultado = CStr(bs.Execute("ValidaProcesso"))

		  If resultado <> "" Then
      	If bsShowMessage(resultado, "Q") = vbNo Then
 			  	Exit Sub
        	End If
      	End If
      End If

      resultado = ""
      bs.ClearParameters
      bs.AddParameter(pdtInteger, HBENEFICIARIO)
      bs.AddParameter(pdtString, SessionVar("EMISSAOCARTAO_ALTERACAOCADASTRAL"))
      bs.AddParameter(pdtString, SessionVar("MIGRACAOCONTRATO"))
      bs.AddParameter(pdtString, CurrentQuery.FieldByName("NAOCOBRAREMISSAO").AsString)
      bs.AddParameter(pdtString, CurrentQuery.FieldByName("IMPRIMIRCARTAO").AsString)
      bs.AddParameter(pdtInteger, CurrentQuery.FieldByName("CARTAOMOTIVO").AsInteger)

      resultado = CStr(bs.Execute("ProcessaEmissaoCartao"))

    End If

	If resultado <> "" Then
    bsShowMessage(resultado, "I")
  End If

  Set bs        = Nothing
End Sub
