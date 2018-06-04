'HASH: E18C8CA4FF78D3BCF3527BA9C1A4032A
'SAM_TIPOGUIA_MDGUIA_CONDATEND
'Rodrigo Soares - 07/04/2006

'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
'Rodrigo Soares - SMS: 60596 - 07/04/2006 - Início
  Dim QTemp As Object
  Set QTemp = NewQuery

  QTemp.Active = False
  QTemp.Clear
  QTemp.Add("SELECT HANDLE")
  QTemp.Add("  FROM SAM_TIPOGUIA_MDGUIA_CONDATEND")
  QTemp.Add(" WHERE CONDICAOATENDIMENTO = :CONDICAOATENDIMENTO")
  QTemp.Add("   AND MODELOGUIA = :MODELOGUIA")
  QTemp.ParamByName("CONDICAOATENDIMENTO").AsInteger = CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger
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

