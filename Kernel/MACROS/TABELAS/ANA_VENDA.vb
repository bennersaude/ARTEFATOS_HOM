'HASH: 87485A56B9FED8BC24E7ABBD7A072A7A
'MACRO  ANA_VENDA
'#Uses "*bsShowMessage"

Option Explicit


Public Sub BOTAOCANCELAR_OnClick()
  Dim interface As Object
  Dim SQL As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT COUNT(HANDLE) QTDE FROM ANA_VENDA_DETALHE WHERE VENDA=" + CurrentQuery.FieldByName("HANDLE").AsString)
  SQL.Active = True

  If SQL.FieldByName("QTDE").AsInteger = 0 Then
    bsShowMessage("A Rotina não está processada !", "I")
    Set SQL = Nothing
    Exit Sub
  End If

  Set interface = CreateBennerObject("SamAnaliseVenda.Rotinas")
  interface.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set SQL = Nothing
  Set interface = Nothing
End Sub

Public Sub BOTAOPLANILHA_OnClick()
  Dim interface As Object
  Dim SQL As Object
  Set SQL = NewQuery

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  SQL.Clear
  SQL.Add("SELECT COUNT(HANDLE) QTDE FROM ANA_VENDA_DETALHE WHERE VENDA=" + CurrentQuery.FieldByName("HANDLE").AsString)
  SQL.Active = True

  If SQL.FieldByName("QTDE").AsInteger = 0 Then
    bsShowMessage("A Rotina não está processada !", "I")
    Set SQL = Nothing
    Exit Sub
  End If

  Set interface = CreateBennerObject("SamAnaliseVenda.Rotinas")
  interface.Planilha(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set SQL = Nothing
  Set interface = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim interface As Object
  Dim SQL As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT COUNT(HANDLE) QTDE FROM ANA_VENDA_DETALHE WHERE VENDA=" + CurrentQuery.FieldByName("HANDLE").AsString)
  SQL.Active = True

  If SQL.FieldByName("QTDE").AsInteger >0 Then
    bsShowMessage("Rotina processada !", "I")
    Set SQL = Nothing
    Exit Sub
  End If


  Set interface = CreateBennerObject("SamAnaliseVenda.Rotinas")
  interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
                      CurrentQuery.FieldByName("DATAFINAL").AsDateTime, CurrentQuery.FieldByName("GRUPOCONTRATO").AsInteger)

  Set SQL = Nothing
  Set interface = Nothing
End Sub

Public Sub TABLE_NewRecord()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Add("SELECT MAX(HANDLE) CODIGO FROM ANA_VENDA")
  SQL.Active = True

  CurrentQuery.FieldByName("CODIGO").AsInteger = SQL.FieldByName("CODIGO").AsInteger + 1
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If(Not CurrentQuery.FieldByName("DATAFINAL").IsNull)And _
     (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)Then
  bsShowMessage("A Data final, se informada, deve ser maior ou igual a inicial", "E")
  CanContinue = False
Else
  CanContinue = True
End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	If (CommandID = "BOTAOCANCELAR") Then
		BOTAOCANCELAR_OnClick
	End If
	If (CommandID = "BOTAOPLANILHA") Then
		BOTAOPLANILHA_OnClick
	End If
	If (CommandID = "BOTAOPROCESSAR") Then
		BOTAOPROCESSAR_OnClick
	End If

End Sub
