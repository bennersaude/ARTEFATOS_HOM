'HASH: C1B39EA38C24C5E9B47B6EDD31E879B0
 
'SAM_TIPOGUIA_MDGUIA_TIPOTRAT
'Rodrigo Soares - 07/04/2006

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
'Rodrigo Soares - SMS: 60596 - 07/04/2006 - Início
  Dim QTemp As Object
  Set QTemp = NewQuery

  QTemp.Active = False
  QTemp.Clear
  QTemp.Add("SELECT HANDLE")
  QTemp.Add("  FROM SAM_TIPOGUIA_MDGUIA_TIPOTRAT")
  QTemp.Add(" WHERE TIPOTRATAMENTO = :TIPOTRATAMENTO")
  QTemp.Add("   AND MODELOGUIA = :MODELOGUIA")
  QTemp.ParamByName("TIPOTRATAMENTO").AsInteger = CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger
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
