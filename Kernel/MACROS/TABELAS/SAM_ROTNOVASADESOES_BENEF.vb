'HASH: E27E3DB9AFEBFA8FC9CC8FA5EE553828

Public Sub TABLE_AfterScroll()
  TableReadOnly = Not CurrentQuery.FieldByName("VALORUNITARIO").IsNull
End Sub
