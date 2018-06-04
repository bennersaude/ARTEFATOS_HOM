'HASH: 1D51FDE4ADEF2024CD6DEF73BD8858B3
'SAM_TIPOGUIA_MDGUIA_LOCALATEND
'Rodrigo Soares - 07/04/2006

'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
'Rodrigo Soares - SMS: 60596 - 07/04/2006 - Início
  Dim QTemp As Object
  Set QTemp = NewQuery

  QTemp.Active = False
  QTemp.Clear
  QTemp.Add("SELECT HANDLE")
  QTemp.Add("  FROM SAM_TIPOGUIA_MDGUIA_LOCALATEND")
  QTemp.Add(" WHERE LOCALATENDIMENTO = :LOCALATENDIMENTO")
  QTemp.Add("   AND MODELOGUIA = :MODELOGUIA")
  QTemp.ParamByName("LOCALATENDIMENTO").AsInteger = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
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


