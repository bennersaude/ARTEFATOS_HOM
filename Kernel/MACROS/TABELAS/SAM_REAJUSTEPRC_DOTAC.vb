'HASH: 6CC463EF26D8BC3C434BB7E5D6A94084

Option Explicit
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"
'#Uses "*NegociacaoPrecos"


Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraTabelaUS(TABELAUS.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELAUS").Value = vHandle
  End If

End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Active = False
  SQL.Add("    SELECT TIPOROTINA FROM SAM_REAJUSTEPRC_PARAM    ")
  SQL.Add("     WHERE HANDLE = :PHANDLE")

  SQL.ParamByName("PHANDLE").AsString = CStr(RecordHandleOfTable("SAM_REAJUSTEPRC_PARAM"))
  SQL.Active = True

  CurrentQuery.FieldByName("TIPOROTINA").Value = SQL.FieldByName("TIPOROTINA").AsInteger

  Set SQL = Nothing
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vAssocociacaodaTabela As Long
  Dim validaNegociacao As String
  Dim vAtedias As Integer
  Dim vDeDias As Integer
  Dim vAteAnos As Integer
  Dim vDeAnos As Integer

  If CurrentQuery.FieldByName("ATEDIAS").IsNull Then
    vAtedias = -1
  Else
    vAtedias = CurrentQuery.FieldByName("ATEDIAS").AsInteger
  End If

  If CurrentQuery.FieldByName("ATEANOS").IsNull Then
    vAteAnos = -1
  Else
  	vAteAnos = CurrentQuery.FieldByName("ATEANOS").AsInteger
  End If

  If CurrentQuery.FieldByName("DEDIAS").IsNull Then
    vDeDias = -1
  Else
    vDeDias = CurrentQuery.FieldByName("DEDIAS").AsInteger
  End If

  If CurrentQuery.FieldByName("DEANOS").IsNull Then
    vDeAnos = -1
  Else
    vDeAnos = CurrentQuery.FieldByName("DEANOS").AsInteger
  End If


  validaNegociacao = ValidarTipoNegociacao(vDeAnos, vDeDias, vAteAnos, vAtedias, CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger)

  If (validaNegociacao <> "") Then
	bsShowMessage(validaNegociacao, "E")
	CanContinue = False
	Exit Sub
  End If

  vAssocociacaodaTabela = CurrentQuery.State
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > _
                              CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then

	bsShowMessage("DATA INICIAL não pode ser maior que a DATA FINAL", "E")

    CanContinue = False
  ElseIf CurrentQuery.FieldByName("NOVAVIGENCIA").AsDateTime <= _
                                    CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then


	bsShowMessage("NOVA VIGÊNCIA deve ser maior que a DATA FINAL", "E")

    CanContinue = False
  End If
End Sub

Public Sub TIPOROTINA_OnChanging(AllowChange As Boolean)
  AllowChange = False
End Sub
