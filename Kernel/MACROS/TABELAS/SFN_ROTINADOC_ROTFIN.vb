'HASH: 570F4F7DD90C30F358DC1D363DE3EB23

'MACRO: SFN_ROTINADOC_ROTFIN
'#Uses "*bsShowMessage"


Option Explicit

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SITUACAO FROM SFN_ROTINADOC WHERE HANDLE=:ROTINADOC")
  SQL.ParamByName("ROTINADOC").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "P" Then
    bsShowMessage("Rotina processada !", "E")
    CanContinue = False
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SITUACAO FROM SFN_ROTINADOC WHERE HANDLE=:ROTINADOC")
  SQL.ParamByName("ROTINADOC").AsInteger = CurrentQuery.FieldByName("ROTINADOC").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "P" Then
    bsShowMessage("Rotina já processada !", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT TABFILTRO, SITUACAO FROM SFN_ROTINADOC WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINADOC")
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "P" Then
    bsShowMessage("Rotina já processada !", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  If SQL.FieldByName("TABFILTRO").AsInteger<>3 Then
    bsShowMessage("A rotina documento deve ter o filtro: Várias rotinas financeiras", "E")
    CanContinue = False
    Set SQL = Nothing
  End If
End Sub

