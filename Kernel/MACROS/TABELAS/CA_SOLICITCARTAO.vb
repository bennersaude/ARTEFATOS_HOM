'HASH: 34928337592FF7B813D1836EC0556403
'Macro: SAM_PRESTADOR
Option Explicit
'############## CENTRAL DE ATENDIMENTO #################

Public Function FU_PodeGerarCartao As Boolean

  Dim SQL As Object


  FU_PodeGerarCartao = False

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT BL.HANDLE FROM SAM_BENEFICIARIO_LICENCA BL ")
  SQL.Add(" WHERE BL.BENEFICIARIO = :HBENEFICIARIO ")
  SQL.Add("   And BL.DATAINICIAL <= :DATA")
  SQL.Add("   And (BL.DATAFINAL Is Null Or BL.DATAFINAL >= :DATA)")

  SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("beneficiario").AsInteger
  SQL.ParamByName("DATA").Value = CurrentQuery.FieldByName("datasolicit").AsDateTime

  SQL.Active = True

  If Not SQL.EOF Then
    MsgBox "Beneficiário em licensa."
    Exit Function
  End If

  SQL.Active = False

  SQL.Clear
  SQL.Add("SELECT BS.HANDLE FROM SAM_BENEFICIARIO_SUSPENSAO BS ")
  SQL.Add(" WHERE BS.BENEFICIARIO = :HBENEFICIARIO")
  SQL.Add("   And BS.DATAINICIAL <= :DATA")
  SQL.Add("   And (BS.DATAFINAL Is Null Or BS.DATAFINAL >= :DATA)")
  SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("beneficiario").AsInteger
  SQL.ParamByName("DATA").Value = CurrentQuery.FieldByName("datasolicit").AsDateTime
  SQL.Active = True

  If Not SQL.EOF Then
    MsgBox "Beneficiário suspenso."
    Exit Function
  End If

  SQL.Active = False

  SQL.Clear
  SQL.Add("SELECT B.HANDLE,B.DATACANCELAMENTO,B.DATABLOQUEIO ")
  SQL.Add("  FROM SAM_BENEFICIARIO b ")
  SQL.Add("WHERE B.handle = :HBENEFICIARIO")
  SQL.Add("  And ((B.DATACANCELAMENTO Is Not Null)")
  SQL.Add("   Or (B.DATABLOQUEIO      Is Not Null))")
  SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("beneficiario").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    MsgBox "Beneficiário Bloqueado/Cancelado."
    Exit Function
  End If

  FU_PodeGerarCartao = True

End Function


Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vCriterios As String
  Dim vCampos As String
  Dim vColunas As String
  Dim Interface As Variant
  Dim vHandle As Long
  Set interface = CreateBennerObject("Procura.Procurar")
  vColunas = "BENEFICIARIO|NOME"
  vCriterios = "DATACANCELAMENTO IS NULL"
  vCampos = "Beneficiário|Nome"
  vHandle = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 2, vCampos, vCriterios, "Tabela de beneficiários", False, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
  ShowPopup = False
  Set Interface = Nothing
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  ' +++++++++dentro da dll da a mensagem de confirmacao
  If MsgBox("Confirma o cancelamento da solicitação?", vbYesNo) = vbNo Then
    Exit Sub
  End If

  Dim vDll As Object
  Dim Retorno As Boolean

  Set vDll = CreateBennerObject("CA016.CARTAO")


  'On Error GoTo erro        "CA_SOLICITCARTAO"

  Retorno = vDll.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  If Retorno = False Then
    Exit Sub
  Else
    WriteAudit("C", HandleOfTable("CA_SOLICITCARTAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Solicitação seg. via de cartão - Cancelamento")
    RefreshNodesWithTable("CA_SOLICITCARTAO")
  End If

  'erro:
  '  MsgBox("Erro na execução da dll CA016")
End Sub

Public Sub TABLE_AfterInsert()
Dim vANO As String
Dim Sequencia As Long
Dim AnoAtual As Long
	If Not VisibleMode Then
		vANO = Format(ServerDate,"yyyy")
		CurrentQuery.FieldByName("ANO").Value = ("01/01/" + vANO)
		AnoAtual = CLng(vANO)
		NewCounter("CA_CARTAO", AnoAtual, 1, Sequencia)
		CurrentQuery.FieldByName("NUMERO").Value = Sequencia
		CurrentQuery.FieldByName("PROTOCOLO").Value = Format(ServerDate,"yyyy") + Format(Sequencia,"######000000")
	End If
End Sub

Public Sub TABLE_AfterPost()
	If Not VisibleMode Then
		InfoDescription = "Por favor anote o número do Protocolo: " + CurrentQuery.FieldByName("PROTOCOLO").Value
	End If
End Sub

Public Sub TABLE_AfterScroll()
  Select Case CurrentQuery.FieldByName("SITUACAO").AsString
    Case "C"
      BOTAOCANCELAR.Visible = False
    Case "P"
      BOTAOCANCELAR.Visible = False
    Case Else
      BOTAOCANCELAR.Visible = True
  End Select
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If FU_PodeGerarCartao = False Then
    Cancontinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_NewRecord()
  Dim vANO As String
  Dim Sequencia As Long
  vANO = Format(ServerDate, "yyyy")
  NewCounter("CA_ATEND", CDate(vANO), 1, Sequencia)
  CurrentQuery.FieldByName("ANO").Value = ("01/01/" + vANO)
  CurrentQuery.FieldByName("NUMERO").Value = Sequencia

  Dim vFilial As Long
  Dim vFilialProcessamento As Long
  Dim vMsg As String
  Dim vNome As String
  Dim vResult As Variant
  vMsg = ""
  vNome = ""
  vResult = BuscarFiliais(CurrentSystem, vFilial, vFilialProcessamento, vMsg)

  If vResult = False Then
    CurrentQuery.FieldByName("filial").Value = vFilial
  Else
    MsgBox "Erro Rotina Buscar filial."
    Exit Sub
  End If
End Sub

'###############################################################
