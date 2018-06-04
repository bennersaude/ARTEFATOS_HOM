'HASH: D0A81CF01A83FBEF963DD6C3D7A86041
'Macro: SAM_PRESTADOR_HORARIO


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("HORAINICIAL").Value > CurrentQuery.FieldByName("HORAFINAL").Value Then
    CanContinue = False
    MsgBox "Hora final nao pode ser menor que a hora inicial!"
    Exit Sub
  End If
  If CurrentQuery.FieldByName("DOMINGO").Value = "N" Then
    If CurrentQuery.FieldByName("SEGUNDA").Value = "N" Then
      If CurrentQuery.FieldByName("TERCA").Value = "N" Then
        If CurrentQuery.FieldByName("QUARTA").Value = "N" Then
          If CurrentQuery.FieldByName("QUINTA").Value = "N" Then
            If CurrentQuery.FieldByName("SEXTA").Value = "N" Then
              If CurrentQuery.FieldByName("SABADO").Value = "N" Then
                CanContinue = False
                MsgBox "Horário inválido - deve-se marcar o(s) dia(s) da semana!"
                Exit Sub
              End If
            End If
          End If
        End If
      End If
    End If
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial ("E", "P", Msg) = "N" Then
    MsgBox Msg
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial ("A", "P", Msg) = "N" Then
    MsgBox Msg
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial ("I", "P", Msg) = "N" Then
    MsgBox Msg
    CanContinue = False
    Exit Sub
  End If
End Sub


