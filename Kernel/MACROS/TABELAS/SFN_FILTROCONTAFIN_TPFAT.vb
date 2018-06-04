'HASH: 590A096B0A02DD3053CC2D3181C98A6D
 
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSel As Object
  Set qSel = NewQuery

  qSel.Active = False
  qSel.Clear
  qSel.Add("SELECT COUNT(HANDLE) QTDE")
  qSel.Add("  FROM SFN_FILTROCONTAFIN_TPFAT")
  qSel.Add(" WHERE TPFATURAMENTO = :TPFATURAMENTO AND HANDLE <> :HANDLE AND FILTROCONTAFIN = :FILTROCONTAFIN")
  qSel.ParamByName("TPFATURAMENTO").AsInteger = CurrentQuery.FieldByName("TPFATURAMENTO").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.ParamByName("FILTROCONTAFIN").AsInteger = CurrentQuery.FieldByName("FILTROCONTAFIN").AsInteger
  qSel.Active = True

  If qSel.FieldByName("QTDE").AsInteger > 0 Then
    MsgBox("Tipo faturamento já cadastrado.")
    CanContinue = False
    Set qSel = Nothing
    Exit Sub
  End If

  Set qSel = Nothing
End Sub
