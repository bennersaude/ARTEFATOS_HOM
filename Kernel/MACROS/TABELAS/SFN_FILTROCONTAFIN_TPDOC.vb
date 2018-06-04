'HASH: 9BE7CB80123ABA5440820621BDDB9783
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSel As Object
  Set qSel = NewQuery

  qSel.Active = False
  qSel.Clear
  qSel.Add("SELECT COUNT(HANDLE) QTDE")
  qSel.Add("  FROM SFN_FILTROCONTAFIN_TPDOC")

  qSel.Add(" WHERE TPDOCUMENTO = :TPDOCUMENTO AND HANDLE <> :HANDLE AND FILTROCONTAFIN = :FILTROCONTAFIN")
  qSel.ParamByName("TPDOCUMENTO").AsInteger = CurrentQuery.FieldByName("TPDOCUMENTO").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.ParamByName("FILTROCONTAFIN").AsInteger = CurrentQuery.FieldByName("FILTROCONTAFIN").AsInteger
  qSel.Active = True

  If qSel.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Tipo de documento já cadastrado.", "E")
    CanContinue = False
    Set qSel = Nothing
    Exit Sub
  End If

  Set qSel = Nothing
End Sub
