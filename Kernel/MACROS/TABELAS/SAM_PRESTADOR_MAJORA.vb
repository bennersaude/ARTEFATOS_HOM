'HASH: 90828021F71BD9AF24991D93C57F5BA3
'Macro: SAM_PRESTADOR_MAJORA

'#Uses "*bsShowMessage"
Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If ((CurrentQuery.FieldByName("DATAFINAL").IsNull) Or (CurrentQuery.FieldByName("DATAFINAL").AsDateTime >= ServerDate)) Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String

  If CurrentQuery.FieldByName("MAJORABASEPF").AsInteger <= 0 And _
                              CurrentQuery.FieldByName("MAJORABASEBONIFICACAO").AsInteger <= 0 Then
    CanContinue = False
    bsShowMessage("Um dos percentuais deve ser maior que zero", "E")
    Exit Sub
  End If

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_MAJORA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  Set Interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  Dim vQuery As BPesquisa

  Set vQuery = NewQuery
  vQuery.Clear
  vQuery.Add("SELECT *")
  vQuery.Add("       FROM SAM_PRESTADOR_MAJORA_EVENTO")
  vQuery.Add("WHERE PRESTADORMAJORA = :PRESTADORMAJORA ")
  vQuery.ParamByName("PRESTADORMAJORA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  vQuery.Active = True

  If Not vQuery.EOF Then
    bsShowMessage("Existem Eventos cadastrados nesta Majoração. Exclusão não permitida", "E")
    CanContinue = False
  End If

  vQuery.Active = False
  Set vQuery = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If


End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

