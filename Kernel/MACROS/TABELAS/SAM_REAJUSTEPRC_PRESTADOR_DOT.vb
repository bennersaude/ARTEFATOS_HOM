'HASH: 56DACD7F0CFA04AFE172AAD6F298C391
'Macro: SAM_REAJUSTEPRC_PRESTADOR_DOT
'#Uses "*bsShowMessage"
'#Uses "*NegociacaoPrecos"

Public Sub TABLE_AfterInsert()
  Dim TIPO As Object
  Set TIPO = NewQuery
  TIPO.Add("SELECT HANDLE FROM SAM_REAJUSTEPRC_PARAMTIPO T WHERE T.REAJUSTEPRCPARAM = :PARAM AND ")
  TIPO.Add("T.TIPODOREAJUSTE = 'D'")
  TIPO.ParamByName("PARAM").Value = CurrentQuery.FieldByName("REAJUSTEPRCPARAM").AsInteger
  TIPO.Active = True
  If TIPO.EOF Then
    SetParamTipo = False
  Else
    setParamTipo = True
    CurrentQuery.FieldByName("PARAMTIPO").Value = TIPO.FieldByName("HANDLE").AsInteger
  End If
  TIPO.Active = False
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem,"E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem,"A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim validaNegociacao As String
  Dim vAtedias As Integer
  Dim vDeDias As Integer
  Dim vAteAnos As Integer
  Dim vDeAnos As Integer

  validaNegociacao = ValidarTipoNegociacao(CurrentQuery.FieldByName("DEANOS").AsInteger, CurrentQuery.FieldByName("DEDIAS").AsInteger, CurrentQuery.FieldByName("ATEANOS").AsInteger, CurrentQuery.FieldByName("ATEDIAS").AsInteger, CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger)

  If (validaNegociacao <> "") Then
	bsShowMessage(validaNegociacao, "E")
	CanContinue = False
	Exit Sub
  End If

  vFiltroAdicional = " AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString

  If Not CurrentQuery.FieldByName("CLASSEASSOCIADO").IsNull Then
    vFiltroAdicional = vFiltroAdicional + " AND CLASSEASSOCIADO = '" + CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString + "'"
  End If

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

  CanContinue = ValidacoesBeforePostNegociacaoPreco(CurrentQuery.FieldByName("HANDLE").AsInteger, _
    "SAM_REAJUSTEPRC_PRESTADOR_DOT", "", "", "PRESTADOR", _
	CurrentQuery.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("EVENTO").AsInteger, "", "-", _
	vFiltroAdicional, vDeAnos, vDeDias, _
	vAteAnos, vAtedias, _
	CurrentQuery.FieldByName("TABNEGOCIACAO").AsInteger, 0, 0)

  If Not CanContinue Then
    Exit Sub
  End If
End Sub
