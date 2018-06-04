'HASH: A4CF33C33ED963A491A5B6D8B5BF32C7
'#Uses "*bsShowMessage"

Option Explicit

Dim simbologia As Integer

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  	Dim qOdontograma As BPesquisa
  	Set qOdontograma = NewQuery

	qOdontograma.Active = False
	qOdontograma.Clear
  	qOdontograma.Add("Select O.HANDLE            ")
  	qOdontograma.Add("  FROM CLI_ODONTOGRAMA O   ")
  	qOdontograma.Add("  Join CLI_INTERVENCAO_TIPO T On O.INTERVENCAOTIPO = T.Handle ")
  	qOdontograma.Add("  Join CLI_INTERVENCAO I On T.INTERVENCAO = I.Handle ")
  	qOdontograma.Add(" WHERE I.HANDLE = :HANDLE  ")
  	qOdontograma.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qOdontograma.Active = True

  	If Not qOdontograma.EOF And CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    	bsShowMessage("Registro não pode ser excluído enquanto existir históricos vinculados.", "E")
    	CanContinue = False
  	End If

  	Set qOdontograma = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qIntervencao As BPesquisa

  Set qIntervencao = NewQuery

  qIntervencao.Active = False
  qIntervencao.Clear
  qIntervencao.Add("SELECT * FROM CLI_INTERVENCAO WHERE CODIGO = :PCODIGO AND HANDLE <> :HANDLE")
  qIntervencao.ParamByName("PCODIGO").AsString = CurrentQuery.FieldByName("CODIGO").AsString
  qIntervencao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qIntervencao.Active = True

  If Not qIntervencao.EOF Then
    bsShowMessage("Código já existente.", "E")
    CanContinue = False
  End If

  Set qIntervencao = Nothing

  If CurrentQuery.FieldByName("SIMBOLOGIAODONTO").AsInteger <> simbologia And simbologia > 0 Then
  	Dim qOdontograma As BPesquisa
  	Set qOdontograma = NewQuery

	qOdontograma.Active = False
	qOdontograma.Clear
  	qOdontograma.Add("Select O.HANDLE            ")
  	qOdontograma.Add("  FROM CLI_ODONTOGRAMA O   ")
  	qOdontograma.Add("  Join CLI_INTERVENCAO_TIPO T On O.INTERVENCAOTIPO = T.Handle ")
  	qOdontograma.Add("  Join CLI_INTERVENCAO I On T.INTERVENCAO = I.Handle ")
  	qOdontograma.Add(" WHERE I.HANDLE = :HANDLE  ")
  	qOdontograma.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qOdontograma.Active = True

  	If Not qOdontograma.EOF And CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
		CurrentQuery.FieldByName("SIMBOLOGIAODONTO").AsInteger = simbologia
    	bsShowMessage("A simbologia não pode ser alterada enquanto existir históricos vinculados.", "E")
    	CanContinue = False
  	End If

  	Set qOdontograma = Nothing
  End If

End Sub

Public Sub TABLE_AfterScroll()

	simbologia = CurrentQuery.FieldByName("SIMBOLOGIAODONTO").AsInteger

End Sub
