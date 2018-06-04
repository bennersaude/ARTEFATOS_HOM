'HASH: 8EDFC92F962C02BD76790249CD186CFD
'MACRO : POR_ARQUIVOCOMPARTILHADO
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("DATAHORAATU").AsDateTime = CurrentSystem.ServerNow
End Sub

Public Sub TABLE_AfterScroll()

  If CurrentQuery.State = 3 Then
    ARQUIVO.ReadOnly = False
    CARREGADOVIA.ReadOnly = False
  Else
    ARQUIVO.ReadOnly = True
    CARREGADOVIA.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If (CurrentQuery.State = 3) Then
    If CurrentQuery.FieldByName("CARREGADOVIA").AsString = "2" Then
		CurrentQuery.FieldByName("CARREGADOVIA").AsString = "1"
    End If

    CurrentQuery.FieldByName("INCLUIDOPOR").AsInteger = CurrentUser
  End If


End Sub
