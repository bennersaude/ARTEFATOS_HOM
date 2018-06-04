'HASH: 890BF958691F8719DFE7B2FAAFBC71BF
 
Public Sub BOTAOGERACODFOLHA_OnClick()
  If CurrentQuery.FieldByName("CODIGOFOLHA").AsString = "I" Then
    Dim interface As Object
    Set interface = CreateBennerObject("SfnGerencial.Rotinas")
    interface.GeraClasses(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,"SFN_FILTROCONTAFIN_CODFOLHA")
    Set interface = Nothing
  End If

End Sub

Public Sub BOTAOGERATPFATURAMENTO_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SfnGerencial.Rotinas")
  interface.GeraClasses(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,"SFN_FILTROCONTAFIN_TPFAT")
  Set interface = Nothing

End Sub

Public Sub BOTAOGERATPLANCFIN_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SfnGerencial.Rotinas")
  interface.GeraClasses(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,"SFN_FILTROCONTAFIN_TPLANCFIN")
  Set interface = Nothing

End Sub

Public Sub BOTAOGERATPDOCUMENTO_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SfnGerencial.Rotinas")
  interface.GeraClasses(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,"SFN_FILTROCONTAFIN_TPDOC")
  Set interface = Nothing
End Sub


Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("CODIGOFOLHA").AsString <> "I" Then
    BOTAOGERACODFOLHA.Enabled= False
  Else
    BOTAOGERACODFOLHA.Enabled= True
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("CODIGOFOLHA").AsString <> "I" Then
    Dim qSel As Object
    Set qSel = NewQuery
    qSel.Active = False
    qSel.Clear
    qSel.Add("SELECT COUNT(HANDLE) QTDE ")
    qSel.Add("  FROM SFN_FILTROCONTAFIN_CODFOLHA")
    qSel.Add(" WHERE FILTROCONTAFIN = :HFILTROCONTAFIN")
    qSel.ParamByName("HFILTROCONTAFIN").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSel.Active = True

    If qSel.FieldByName("QTDE").AsInteger > 0 Then
      MsgBox("Não é possível alterar o filtro de código folha pois existe código folha cadastrado para o filtro da ficha financeira.")
      CanContinue = False
      Set qSel = Nothing
      Exit Sub
    End If

    Set qSel = Nothing
  End If

End Sub
