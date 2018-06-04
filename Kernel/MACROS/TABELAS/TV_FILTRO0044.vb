'HASH: 407C18F47563A7BC3B25AE01E1094219
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If VerificarEvento Then
	bsShowMessage("Evento inicial maior que o evento final !", "E")
	CanContinue = False
  End If

  If (Not(BuscarExisteReembolsoContrato) And Not(BuscarExisteReembolsoDotacao)) Then
    bsShowMessage("O contrato do beneficiário não possui cadastro para preço por reembolso!", "E")
    CanContinue = False
  End If
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOINICIAL.Text)

  If vHandle <> 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOFINAL.Text)

  If vHandle <> 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
End Sub

' Verificar Evento inicial e final
Public Function VerificarEvento As Boolean
  Dim vData As String
  Dim sEventoInicial As String
  Dim sEventoFinal As String
  Dim Q As Object

  vData = SQLDate(ServerDate)
  sEventoInicial = CurrentQuery.FieldByName("EVENTOINICIAL").AsString
  sEventoFinal  = CurrentQuery.FieldByName("EVENTOFINAL").AsString

  Set Q = NewQuery
  Q.Clear
  Q.Add("SELECT E1.ESTRUTURA ESTRUTURAINICIAL, E2.ESTRUTURA ESTRUTURAFINAL ")
  Q.Add("  FROM SAM_TGE E1,                                                ")
  Q.Add("       SAM_TGE E2                                                 ")
  Q.Add(" WHERE E1.HANDLE = " + sEventoInicial								)
  Q.Add("   AND E2.HANDLE = " + sEventoFinal								)
  Q.Active = True

  VerificarEvento = (Q.FieldByName("ESTRUTURAINICIAL").AsString > Q.FieldByName("ESTRUTURAFINAL").AsString)
  Set Q = Nothing
End Function

' Verificar se existe configuração de reembolso por contrato
Public Function BuscarExisteReembolsoContrato As Boolean
  Dim vData As String
  Dim sBeneficiario As String
  Dim Q As Object

  vData = SQLDate(ServerDate)
  sBeneficiario = CurrentQuery.FieldByName("BENEFICIARIO").AsString

  Set Q = NewQuery
  Q.Clear
  Q.Add("SELECT HANDLE                      ")
  Q.Add("  FROM SAM_CONTRATO_PRECOREEMBOLSO ")
  Q.Add(" WHERE DATAINICIAL <= " + vData)
  Q.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= " + vData +")")
  Q.Add("   AND CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = " + sBeneficiario +")")
  Q.Active = True

  BuscarExisteReembolsoContrato = Not(Q.EOF)
  Set Q = Nothing
End Function

' Verificar se existe configuração de reembolso por Dotação
Public Function BuscarExisteReembolsoDotacao As Boolean

  Dim sBeneficiario As String
  Dim Q As Object

  sBeneficiario = CurrentQuery.FieldByName("BENEFICIARIO").AsString

  Set Q = NewQuery
  Q.Clear
  Q.Add("SELECT REG.HANDLE 										")
  Q.Add("  FROM SAM_CONTRATO_PRECOREEMBDOT_REG REG 		")
  Q.Add(" where  DATAINICIAL <= :DATA						 	")
  Q.Add("   AND (DATAFINAL IS NULL OR DATAFINAL <= :DATA)		")
  Q.Add("   AND CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO) ")
  Q.Add("union													")
  Q.Add("SELECT MUN.HANDLE 										")
  Q.Add("  FROM SAM_CONTRATO_PRECOREEMBDOT_MUN MUN 		")
  Q.Add(" where  DATAINICIAL <= :DATA						 	")
  Q.Add("   AND (DATAFINAL IS NULL OR DATAFINAL <= :DATA)		")
  Q.Add("   AND CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO) ")
  Q.Add("union 													")
  Q.Add("SELECT EST.HANDLE 										")
  Q.Add("  FROM SAM_CONTRATO_PRECOREEMBDOT_EST EST 		")
  Q.Add(" where  DATAINICIAL <= :DATA						 	")
  Q.Add("   AND (DATAFINAL IS NULL OR DATAFINAL <= :DATA)		")
  Q.Add("   AND CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO) ")
  Q.Add("union 													")
  Q.Add("SELECT GER.HANDLE 										")
  Q.Add("  FROM SAM_CONTRATO_PRECOREEMBDOT_GER GER 		")
  Q.Add(" where  DATAINICIAL <= :DATA						 	")
  Q.Add("   AND (DATAFINAL IS NULL OR DATAFINAL <= :DATA)		")
  Q.Add("   AND CONTRATO = (SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO) ")
  Q.ParamByName("DATA").Value = ServerDate
  Q.ParamByName("BENEFICIARIO").Value = sBeneficiario
  Q.Active = True

  BuscarExisteReembolsoDotacao = Not(Q.EOF)
  Set Q = Nothing
End Function
