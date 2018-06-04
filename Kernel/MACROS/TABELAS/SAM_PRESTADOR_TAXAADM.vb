'HASH: 65CC21750A5B04DC0156598430D377E5
'MACRO: SAM_PRESTADOR_TAXAADM
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Query As Object

  ' Leonam - SMS 36033 - valida data final sendo maior ou igual a data inicial
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
      CanContinue = False

      bsShowMessage("Data final da vigência não pode ser menor que a data inicial!", "E")
      Exit Sub
    End If
  End If

  Set Query = NewQuery

  ' Leonam - SMS 36033 - valida data inicial maior ou igual a data de inclusao do prestador
  Query.Active = False
  Query.Clear
  Query.Add("SELECT DATAINCLUSAO        ")
  Query.Add("  FROM SAM_PRESTADOR       ")
  Query.Add(" WHERE HANDLE = :PRESTADOR ")
  Query.Add("   AND DATAINCLUSAO IS NOT NULL    ")
  Query.Add("   AND DATAINCLUSAO > :DATAINICIAL ")
  Query.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  Query.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  Query.Active = True

  If Not Query.FieldByName("DATAINCLUSAO").IsNull Then
    CanContinue = False
    bsShowMessage("Data inicial da vigência não pode ser menor que a data de inclusão do prestador!", "E")

    Query.Active = False
    Set Query = Nothing

    Exit Sub
  End If


  ' Leonam - SMS 36033 - consulta referência cruzada
  Query.Active = False
  Query.Clear
  Query.Add("SELECT COUNT(HANDLE) QNT     ")
  Query.Add("  FROM SAM_PRESTADOR_TAXAADM ")
  Query.Add(" WHERE HANDLE <> :HANDLE     ")
  Query.Add("   AND PRESTADOR = :PRESTADOR")
  Query.Add("   AND ((DATAFINAL >= :DATAINICIAL OR DATAFINAL IS NULL) AND (DATAINICIAL <= :DATAFINAL OR :DATAFINAL IS NULL)) ")

  If SessionVar("TAXAADMPRESTADORPOREMPRESA") <> "" Then
    Query.Add(SessionVar("TAXAADMPRESTADORPOREMPRESA"))
  End If

  Query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  Query.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  Query.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime

  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    Query.ParamByName("DATAFINAL").DataType = ftDateTime
    Query.ParamByName("DATAFINAL").Clear
  Else
    Query.ParamByName("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  End If

  Query.Active = True

  If Query.FieldByName("QNT").AsInteger > 0 Then
    CanContinue = False

    bsShowMessage("A vigência informada esta cruzando com outra já cadastrada!", "E")
  End If

  Set Query = Nothing
End Sub

