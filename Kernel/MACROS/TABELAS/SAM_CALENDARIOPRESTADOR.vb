﻿'HASH: 4CAA4C543B1592F6B8CB6492DC8FFF2F
'Macro: SAM_CALENDARIOPRESTADOR

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CALENDARIOPRESTADOR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub


Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  '  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
  '    CanContinue =False
  '    MsgBox "Operação cancelada. Usuário diferente"
  '    Exit Sub
  '  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

