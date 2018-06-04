'HASH: C36226BD175BEADED280C7A2D9994F58

'CLI_ROTREMARCAFALTAAGENDA

Public Sub BOTAOREMARCAR_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    Dim BSCli001dll As Object
    Set BSCli001dll = CreateBennerObject("BSCli001.Rotinas")
    BSCli001dll.RemarcaFaltaRecurso(CurrentSystem, _
                                    CurrentQuery.FieldByName("AGENDA").AsInteger, _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger)
    RefreshNodesWithTable("CLI_ROTREMARCAFALTAAGENDA")
    Set BSCli001dll = Nothing
  Else
    MsgBox("Para remarcar a situação deve estar pendente!")
  End If
End Sub

