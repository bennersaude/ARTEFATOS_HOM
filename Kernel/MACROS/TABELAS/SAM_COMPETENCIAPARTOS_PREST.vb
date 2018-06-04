'HASH: D8113F9DFDB9D49C30145C5E22A2C4A8
'Macro: SAM_COMPETENCIAPARTOS_PREST
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
   Dim query As BPesquisa
   Set query = NewQuery
   query.Active = False
   query.Clear
   query.Add("SELECT *                             ")
   query.Add("  FROM SAM_COMPETENCIAPARTOS_PREST   ")
   query.Add("  WHERE PRESTADOR= :PHANDLEPRESTADOR ")
   query.Add("    AND COMPETENCIA= :PCOMPETENCIA   ")
   query.Add("    AND HANDLE<> :PHANDLE            ")
   query.ParamByName("PHANDLEPRESTADOR").AsInteger= CurrentQuery.FieldByName("PRESTADOR").AsInteger
   query.ParamByName("PCOMPETENCIA").AsInteger = CurrentQuery.FieldByName("COMPETENCIA").AsInteger
   query.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
   query.Active = True

   If Not query.EOF Then
     bsShowMessage("Existe uma competência informada para este prestador!", "I")
     CanContinue=False
   End If
  Set query=Nothing
End Sub
