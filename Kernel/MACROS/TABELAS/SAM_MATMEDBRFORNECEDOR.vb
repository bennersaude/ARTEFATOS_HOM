'HASH: 625EDB1481ECCDD45803605A8E1740FB
Public Sub TABLE_NewRecord()

   Dim sql As BPesquisa
   Set sql = NewQuery

   sql.Active = False
   sql.Clear
   sql.Add("SELECT MAX(CODIGO) CODIGO     ")
   sql.Add("  FROM SAM_MATMEDBRFORNECEDOR ")
   sql.Active = True

   CurrentQuery.FieldByName("CODIGO").AsInteger = sql.FieldByName("CODIGO").AsInteger + 1

   Set sql = Nothing

End Sub
