'HASH: 3E5F7BE0A3878EA373B51CEE44279E50

' Macro : SAM_PEGARQUIVOFISICO

Public Sub BOTAOINCLUIRPEG_OnClick()
  If ((CurrentQuery.State <> 2) And (CurrentQuery.State <> 3)) Then
    If (CurrentQuery.FieldByName("USUARIOINICIAL").IsNull) Then
      Dim qUpdateArquivo As Object
      Set qUpdateArquivo = NewQuery
      qUpdateArquivo.Active = False
      qUpdateArquivo.Clear
      qUpdateArquivo.Add("UPDATE SAM_PEGARQUIVOFISICO SET USUARIOINICIAL = :pUSUARIO, DATAHORAINICIAL = :pDATA WHERE HANDLE = :pHANDLE")
      qUpdateArquivo.ParamByName("pUSUARIO").AsInteger = CurrentUser
      qUpdateArquivo.ParamByName("pDATA").AsDateTime = CurrentSystem.ServerNow
      qUpdateArquivo.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qUpdateArquivo.ExecSQL
      Set qUpdateArquivo = Nothing
    End If
  Else
    MsgBox("O registro não pode estar em edição.")
    Exit Sub
  End If
  Dim dllPegArquivo As Object
  Set dllPegArquivo = CreateBennerObject("BSPEGARQUIVO.ROTINAS")
  dllPegArquivo.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set dllPegArquivo = Nothing
  RefreshNodesWithTable("SAM_PEGARQUIVOFISICO")
End Sub

Public Sub TABLE_AfterInsert()
  Dim vCodigo As Long
  ' Atribui ao código, o valor do contador
  vCodigo = -1
  CurrentSystem.NewCounter("SAM_PEGARQUIVOFISICO", 0, 1, vCodigo)
  CurrentQuery.FieldByName("CODIGO").AsInteger = vCodigo
End Sub

Public Sub TABLE_AfterPost()
  BOTAOINCLUIRPEG_OnClick
End Sub

