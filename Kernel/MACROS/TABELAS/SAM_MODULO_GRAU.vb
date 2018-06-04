'HASH: D3C3FC972E6B4C557963BCDBA3F1571F
'#Uses "*bsShowMessage"

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vColunas As String
  Dim vCriterios As String
  Dim vHandle As Long
  Dim vCampos As String
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_GRAU.GRAU|SAM_GRAU.DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  vCriterios = "HANDLE > 0 And TIPOGRAU In (Select HANDLE FROM SAM_TIPOGRAU WHERE CLASSIFICACAO = '3')"
  vCampos = "Grau|Descrição|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterios, "Tabela de Graus", True, "")
  CurrentQuery.FieldByName("GRAU").AsInteger = vHandle
  Set interface = Nothing
  ShowPopup = False


End Sub



Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qVerificaDuplicidade As Object
  Set qVerificaDuplicidade = NewQuery
  qVerificaDuplicidade.Active = False
  qVerificaDuplicidade.Clear
  qVerificaDuplicidade.Add("SELECT Count(1) Encontrou ")
  qVerificaDuplicidade.Add("  FROM SAM_MODULO_GRAU ")
  qVerificaDuplicidade.Add(" WHERE HANDLE <> "+CurrentQuery.FieldByName("HANDLE").AsString  )
  qVerificaDuplicidade.Add("   AND GRAU =  "+CurrentQuery.FieldByName("GRAU").AsString )
  qVerificaDuplicidade.Add("   AND MODULO = "+CurrentQuery.FieldByName("MODULO").AsString )
  qVerificaDuplicidade.Active = True

  If qVerificaDuplicidade.FieldByName("Encontrou").AsInteger > 0 Then
    bsShowMessage("Grau para acomodação já cadastrado.", "E")
    CanContinue = False
    Set qVerificaDuplicidade = Nothing
    Exit Sub
  End If

  Set qVerificaDuplicidade = Nothing

End Sub

Public Sub TABLE_BeforeScroll()
	If WebMode Then
		GRAU.WebLocalWhere =  "HANDLE > 0 AND TIPOGRAU IN (SELECT HANDLE FROM SAM_TIPOGRAU WHERE CLASSIFICACAO = '3')"
	End If
End Sub
