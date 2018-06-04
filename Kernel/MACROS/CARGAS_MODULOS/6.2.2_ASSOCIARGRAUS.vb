'HASH: A68534769DBD5BBEC97A640FA7957B11
 
Public Sub GRAUSODONTOLOGICOS_OnClick()
Dim SQL As Object
Set SQL =NewQuery
  SQL.Clear
  SQL.Add("SELECT 1")
  SQL.Add("  FROM SAM_TGE")
  SQL.Add(" WHERE HANDLE = :HANDLE ")
  SQL.Add("   AND TABTIPOEVENTO = 2")
  SQL.ParamByName("HANDLE").AsInteger =RecordHandleOfTable("SAM_TGE")
  SQL.Active =True
  If SQL.EOF Then
    MsgBox("Este evento não foi definido como odontológico!")
    Exit Sub
  End If
  Dim ODONTO As Object
  Set ODONTO =CreateBennerObject("BSCLI006.ROTINAS")
  ODONTO.ASSOCIARGRAUS(CurrentSystem,RecordHandleOfTable("SAM_TGE"))

  Set ODONTO =Nothing
  Set SQL =Nothing
End Sub
