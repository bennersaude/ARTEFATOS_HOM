'HASH: ACAD25B7A15E40AF1801323A0314574E
 
'SAM_TIPOGUIA_MDGUIA_REGATEND
'Rodrigo Soares - 07/04/2006

'#uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
'Rodrigo Soares - SMS: 60596 - 07/04/2006 - Início
  Dim QTemp As Object
  Set QTemp = NewQuery

  QTemp.Active = False
  QTemp.Clear
  QTemp.Add("SELECT HANDLE")
  QTemp.Add("  FROM SAM_TIPOGUIA_MDGUIA_REGATEND")
  QTemp.Add(" WHERE REGIMEATENDIMENTO = :REGIMEATENDIMENTO")
  QTemp.Add("   AND MODELOGUIA = :MODELOGUIA")
  QTemp.ParamByName("REGIMEATENDIMENTO").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
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
