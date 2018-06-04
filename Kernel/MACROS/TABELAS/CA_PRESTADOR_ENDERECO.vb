'HASH: A5F3BEA77EFA476465BA36D3A64F6E25
'Macro: CA_PRESTADOR_ENDERECO
'#Uses "*bsShowMessage"
'#Uses "*checkPermissaoFilial"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("CORRESPONDENCIA").Value = "S" Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CORRESPONDENCIA FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = :PRESTADOR")
    SQL.Add("AND CORRESPONDENCIA = 'S' AND HANDLE <> :HCORRENTE")
    SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
    SQL.ParamByName("HCORRENTE").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.Active = True
    If Not SQL.EOF Then
      CurrentQuery.FieldByName("CORRESPONDENCIA").Value = "N"
      CanContinue = False
      bsShowMessage("Existe outro endereco marcado para correspondência!", "E")
    End If
    Set SQL = Nothing
  End If
  If CurrentQuery.FieldByName("CORRESPONDENCIA").Value = "N" And CurrentQuery.FieldByName("ATENDIMENTO").Value = "N" Then
    bsShowMessage("Endereço deve ser marcado para correspondência ou atendimento!", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial("E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial ("A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial ("I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
