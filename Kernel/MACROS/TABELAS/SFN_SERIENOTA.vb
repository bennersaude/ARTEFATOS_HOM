'HASH: 9217768DD2E7C746C6BEB433DAC4F700
'Incluído por Keila em 10-04-2002
'SMS: 7814
'#Uses "*bsShowMessage"


Public Sub TABLE_AfterPost()
  ULTIMONUMERO.ReadOnly = True
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim qAux As Object
	Set qAux = NewQuery
	qAux.Add("SELECT HANDLE        ")
    qAux.Add("  FROM SFN_TIPONOTA  ")
    qAux.Add(" WHERE SERIE = :SERIE")
    qAux.Add(" UNION               ")
    qAux.Add("SELECT HANDLE        ")
    qAux.Add("  FROM SFN_NOTA      ")
    qAux.Add(" WHERE SERIE = :SERIE")
	qAux.ParamByName("SERIE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qAux.Active = True

	If Not qAux.EOF Then
		CanContinue = False
		bsShowMessage("Não foi possível excluir", "E")
	End If

	Set qAux = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Add("SELECT HANDLE FROM SFN_NOTA WHERE NUMERO = " + CurrentQuery.FieldByName("ULTIMONUMERO").AsString)
  Sql.Active = True

  If Sql.EOF Then
    ULTIMONUMERO.ReadOnly = False
  End If
  Set Sql = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  ULTIMONUMERO.ReadOnly = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Sql As Object
  Dim qVerSerie As Object
  Set Sql = NewQuery
  Set qVerSerie = NewQuery

  CanContinue = True

  Sql.Add("SELECT HANDLE FROM SFN_NOTA WHERE NUMERO =" + CurrentQuery.FieldByName("ULTIMONUMERO").AsString)
  Sql.Active = True

  If Not Sql.EOF Then
    bsShowMessage("Já existe Nota Fiscal com esse Número", "E")
    CanContinue = False
    Exit Sub
  End If

  qVerSerie.Add("SELECT HANDLE FROM SFN_SERIENOTA")
  qVerSerie.Add(" WHERE SERIE = '" + CurrentQuery.FieldByName("SERIE").AsString + "'")
  qVerSerie.Add("   AND ULTIMONUMERO = " + CurrentQuery.FieldByName("ULTIMONUMERO").AsString)
  If CurrentQuery.FieldByName("HANDLE").AsInteger <> -1 Then
    qVerSerie.Add("   AND HANDLE <> " + CurrentQuery.FieldByName("HANDLE").AsString)
  End If
  qVerSerie.Active = True

  If Not qVerSerie.EOF Then
	bsShowMessage("Já existe essa série cadastrada", "E")
	CanContinue = False
  End If

  Set Sql = Nothing
  Set qVerSerie = Nothing
End Sub

Public Sub TABLE_BeforeScroll()
	If VisibleMode Then
  		ULTIMONUMERO.ReadOnly = True
  	End If
End Sub
