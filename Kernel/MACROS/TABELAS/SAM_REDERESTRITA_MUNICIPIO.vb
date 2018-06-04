'HASH: F335649B5578EC4A3389B66DF222D70B
'Macro: SAM_REDERESTRITA_MUNICIPIO
'#Uses "*bsShowMessage"


Public Sub BOTAODUPLICAR_OnClick()
  Dim DuplicaRedeRestritaDLL As Object
  Set DuplicaRedeRestritaDLL = CreateBennerObject("SamDupRedeRestrita.SamDupRedeRestrita")
  DuplicaRedeRestritaDLL.Executar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set DuplicaRedeRestritaDLL = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "E", "P", Msg)
  If vFitro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "A", "P", Msg)
  If vFiltro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "I", "P", Msg)
  If vFiltro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAODUPLICAR") Then
		BOTAODUPLICAR_OnClick
	End If
End Sub
