'HASH: 63EFCB6126C8ED62249254E6C21F3577
 
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSel As Object
  Set qSel = NewQuery

  qSel.Active = False
  qSel.Clear
  qSel.Add("SELECT COUNT(HANDLE) QTDE")
  qSel.Add("  FROM SFN_FILTROCONTAFIN_CODFOLHA")
  qSel.Add(" WHERE CODIGOFOLHA = :CODIGOFOLHA AND HANDLE <> :HANDLE AND FILTROCONTAFIN = :FILTROCONTAFIN")
  qSel.ParamByName("CODIGOFOLHA").AsInteger = CurrentQuery.FieldByName("CODIGOFOLHA").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.ParamByName("FILTROCONTAFIN").AsInteger = CurrentQuery.FieldByName("FILTROCONTAFIN").AsInteger
  qSel.Active = True

  If qSel.FieldByName("QTDE").AsInteger > 0 Then
    MsgBox("Código folha já cadastrado.")
    CanContinue = False
    Set qSel = Nothing
    Exit Sub
  End If

  Set qSel = Nothing

End Sub
