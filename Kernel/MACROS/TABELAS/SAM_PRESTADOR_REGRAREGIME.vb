'HASH: E81E13F58DD0A78874D6D7FA36F197D4
'Macro: SAM_PRESTADOR_REGRA
'02/01/2001 - Alterado por Paulo Garcia Junior - liberacao para edição do registro atraves dos parametros gerais de prestador
'#Uses "*liberaRegraExcecao"
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
  If liberaRegraExcecao <> "" Then
    PRESTADOR.ReadOnly = True
    REGRA.ReadOnly = True
  Else
    PRESTADOR.ReadOnly = False
    REGRA.ReadOnly = False
  End If


End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaRegraExcecao
  If Msg<>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaRegraExcecao
  If Msg<>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaRegraExcecao
  If Msg<>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If
End Sub

