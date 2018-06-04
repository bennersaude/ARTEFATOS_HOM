'HASH: 1E025228E321F4F34E9E3BC2D49E60CA
Dim sMacro As String

Public Sub BTNMACRO_OnClick()
  Dim obj As Object, Nome As String, iTIPO As Integer
  Nome = CurrentQuery.FieldByName("FUNCAO").AsString
  sMacro = CurrentQuery.FieldByName("MACRO").AsString
  iTIPO = CurrentQuery.FieldByName("TIPO").AsInteger
  Set obj = CreateBennerObject("Calc.BCalc")
  sMacro = obj.ExecMacroFunction(Nome, sMacro, iTIPO)
  Set obj = Nothing
  CurrentQuery.Edit
  On Error GoTo runnervelhoBTN
  If ExternalMacro Then
    If Trim(sMacro) <> "" Then
      CurrentQuery.FieldByName("MACRO").Value = "'Novo engenho de macro"
    Else
      CurrentQuery.FieldByName("MACRO").Clear
    End If
  Else
    CurrentQuery.FieldByName("MACRO").Value = sMacro
  End If
  Exit Sub
runnervelhoBTN:
  CurrentQuery.FieldByName("MACRO").Value = sMacro
End Sub


Public Sub TABLE_AfterDelete()
  On Error GoTo runnervelhoDEL
  If ExternalMacro Then
    DeleteMacroFromBDoc(mkCalcFunc, CurrentQuery.FieldByName("HANDLE").AsInteger)
  End If
runnervelhoDEL:
End Sub


Public Sub TABLE_AfterPost()
  On Error GoTo runnervelhoPOST
  If ExternalMacro Then
    SaveMacroToBDoc(mkCalcFunc, CurrentQuery.FieldByName("HANDLE").AsInteger, sMacro)
  End If
runnervelhoPOST:
End Sub


Public Sub TABLE_AfterScroll()
  MACRO.Visible = False
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("TIPO").AsInteger = 2) And (CurrentQuery.FieldByName("MACRO").AsString = "") Then
    If VisibleMode Then
      MsgBox("Campo ""MACRO"" é obrigatório")
    Else
      CancelDescription = "Campo ""MACRO"" é obrigatório"
    End If
    CanContinue = False
  End If
End Sub

