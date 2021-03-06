﻿'HASH: E0F00F2727D6FCEDC790D6B9E16423EA
Option Explicit

'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"
'#Uses "*NegociacaoPrecos"

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraTabelaUS(TABELAUS.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELAUS").Value = vHandle
  End If

End Sub

Public Sub TABLE_AfterPost()
  RefreshNodesWithTable("SAM_REAJUSTEPRC_PARAMREGIME")
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim validaNegociacao As String
  Dim vAtedias As Integer
  Dim vDeDias As Integer
  Dim vAteAnos As Integer
  Dim vDeAnos As Integer

  If CurrentQuery.FieldByName("ATEDIAS").IsNull Then
    vAtedias = -1
  Else
    vAtedias = CurrentQuery.FieldByName("ATEDIAS").AsInteger
  End If

  If CurrentQuery.FieldByName("ATEANOS").IsNull Then
    vAteAnos = -1
  Else
  	vAteAnos = CurrentQuery.FieldByName("ATEANOS").AsInteger
  End If

  If CurrentQuery.FieldByName("DEDIAS").IsNull Then
    vDeDias = -1
  Else
    vDeDias = CurrentQuery.FieldByName("DEDIAS").AsInteger
  End If

  If CurrentQuery.FieldByName("DEANOS").IsNull Then
    vDeAnos = -1
  Else
    vDeAnos = CurrentQuery.FieldByName("DEANOS").AsInteger
  End If

  validaNegociacao = ValidarTipoNegociacao(vDeAnos, vDeDias, vAteAnos, vAtedias, CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger)

  If (validaNegociacao <> "") Then
	bsShowMessage(validaNegociacao, "E")
	CanContinue = False
	Exit Sub
  End If

  Condicao = ""

  Set interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = interface.Vigencia(CurrentSystem, "SAM_REAJUSTEPRC_GRAU", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "HANDLE", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > _
                              CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    bsShowMessage("Data INICIAL não pode ser maior que a data FINAL", "E")
    CanContinue = False
  ElseIf CurrentQuery.FieldByName("NOVAVIGENCIA").AsDateTime <= _
                                    CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    bsShowMessage("NOVA VIGÊNCIA deve ser maior que a data FINAL", "E")
    CanContinue = False
  End If
End Sub

