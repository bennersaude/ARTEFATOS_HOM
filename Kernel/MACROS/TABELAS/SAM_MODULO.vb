'HASH: EFE9424A6FFB5FA2695183B605B5C490
'#uses "*bsShowMessage"

Public Sub BOTAODUPLICARMODULO_OnClick()
  Dim DuplicaModuloDLL As Object
  Dim mensagem As String

  If VisibleMode Then

	Set DuplicaModuloDLL = CreateBennerObject("BSINTERFACE0009.Rotinas")
   	DuplicaModuloDLL.GERAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Obj = Nothing
  Else

    Set DuplicaModuloDLL = CreateBennerObject("SamDuplicModulo.DuplicarModulo_Gerar")

    If (DuplicaModuloDLL.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, mensagem) > 0 ) Then
       bsShowMessage(mensagem, "I")
    ElseIf (Len(mensagem)) >0 Then
       bsShowMessage(mensagem, "I")
    End If
    Set DuplicaModuloDLL = Nothing
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If VisibleMode Then
    BOTAOGERAREVENTOS.Visible =False
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAODUPLICARMODULO" Then
		BOTAODUPLICARMODULO_OnClick
	End If
	If CommandID = "DuplicarModulo" Then
		BOTAODUPLICARMODULO_OnClick
	End If
End Sub
