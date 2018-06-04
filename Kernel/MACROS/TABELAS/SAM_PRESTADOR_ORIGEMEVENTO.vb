'HASH: 315CAF6BE903676A2595269286C384C7
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "" '"AND ORIGEMEVENTO = " + CurrentQuery.FieldByName("ORIGEMEVENTO").AsString

  If VisibleMode = True Then
    Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ORIGEMEVENTO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)
    Set Interface = Nothing

    If (Linha <> "") Then
      CanContinue = False
      bsShowMessage(Linha, "E")
      Exit Sub
    End If
  End If

  If (CurrentQuery.FieldByName("DATAULTIMOACESSO").IsNull) Then
    CurrentQuery.FieldByName("DATAULTIMOACESSO").AsDateTime = ServerNow
  End If

End Sub

