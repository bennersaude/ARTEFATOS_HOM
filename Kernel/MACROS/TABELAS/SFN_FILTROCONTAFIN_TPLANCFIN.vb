'HASH: 623B2037C416540EEFE9684468C5686D
 
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSel As Object
  Set qSel = NewQuery

  qSel.Active = False
  qSel.Clear
  qSel.Add("SELECT COUNT(HANDLE) QTDE")
  qSel.Add("  FROM SFN_FILTROCONTAFIN_TPLANCFIN")
  qSel.Add(" WHERE TIPOLANCFIN = :TIPOLANCFIN AND HANDLE <> :HANDLE AND FILTROCONTAFIN = :FILTROCONTAFIN")
  qSel.ParamByName("TIPOLANCFIN").AsInteger = CurrentQuery.FieldByName("TIPOLANCFIN").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.ParamByName("FILTROCONTAFIN").AsInteger = CurrentQuery.FieldByName("FILTROCONTAFIN").AsInteger
  qSel.Active = True

  If qSel.FieldByName("QTDE").AsInteger > 0 Then
    MsgBox("Tipo lançamento financeiro já cadastrado.")
    CanContinue = False
    Set qSel = Nothing
    Exit Sub
  End If

  Set qSel = Nothing

End Sub
