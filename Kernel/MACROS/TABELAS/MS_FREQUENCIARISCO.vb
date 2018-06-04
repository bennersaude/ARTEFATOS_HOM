'HASH: EA4C107A61013423B811807B88D37FAF
'MACRO: MS_FREQUENCIARISCO

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim VERIFICA As Object
  Set VERIFICA = NewQuery

  VERIFICA.Active = False
  VERIFICA.Clear
  VERIFICA.Add("SELECT HANDLE            ")
  VERIFICA.Add("  FROM MS_FREQUENCIARISCO")
  VERIFICA.Add(" WHERE HANDLE <> :HANDLE ")
  VERIFICA.Add("   AND CODIGO = :CODIGO  ")
  VERIFICA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  VERIFICA.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  VERIFICA.Active = True

  If Not VERIFICA.EOF Then
    MsgBox "Já existe um registro para este código!"
    CanContinue = False
  End If

  Set VERIFICA = Nothing
End Sub
 
