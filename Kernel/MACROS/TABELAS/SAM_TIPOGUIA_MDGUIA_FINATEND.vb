'HASH: 763832DA4E2419A3FF61C16BFAA7F8E1
'SAM_TIPOGUIA_MDGUIA_FINATEND
'Rodrigo Soares - 07/04/2006

'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
'Rodrigo Soares - SMS: 60596 - 07/04/2006 - Início
  Dim QTemp As Object
  Set QTemp = NewQuery

  QTemp.Active = False
  QTemp.Clear
  QTemp.Add("SELECT HANDLE")
  QTemp.Add("  FROM SAM_TIPOGUIA_MDGUIA_FINATEND")
  QTemp.Add(" WHERE FINALIDADEATENDIMENTO = :FINALIDADEATENDIMENTO")
  QTemp.Add("   AND MODELOGUIA = :MODELOGUIA")
  QTemp.ParamByName("FINALIDADEATENDIMENTO").AsInteger = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
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
