'HASH: DCFC1ABD84496E8D8649320343DD3F5E
'SAM_TIPOGUIA_MDGUIA_OBJTRAT
'Rodrigo Soares - 07/04/2006

'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
'Rodrigo Soares - SMS:`60596 - 07/04/2006 - Início
  Dim QTemp As Object
  Set QTemp = NewQuery

  QTemp.Active = False
  QTemp.Clear
  QTemp.Add("SELECT HANDLE")
  QTemp.Add("  FROM SAM_TIPOGUIA_MDGUIA_OBJTRAT")
  QTemp.Add(" WHERE OBJETIVOTRATAMENTO = :OBJETIVOTRATAMENTO")
  QTemp.Add("   AND MODELOGUIA = :MODELOGUIA")
  QTemp.ParamByName("OBJETIVOTRATAMENTO").AsInteger = CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger
  QTemp.ParamByName("MODELOGUIA").AsInteger = CurrentQuery.FieldByName("MODELOGUIA").AsInteger
  QTemp.Active = True
  If Not QTemp.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Registro já cadastrado!", "E")
    CanContinue = False
    Set QTemp = Nothing
    Exit Sub
  End If
  Set QTemp = Nothing
  'Rodrigo Soares - SMS: 60596 - 07/04/2006 - Fim
End Sub
