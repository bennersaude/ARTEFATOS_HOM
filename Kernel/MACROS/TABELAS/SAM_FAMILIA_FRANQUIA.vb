'HASH: 5566DBCC35D1BA3B2364776BAC418B4A
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = " AND FRANQUIA = " + CurrentQuery.FieldByName("FRANQUIA").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_FAMILIA_FRANQUIA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FAMILIA", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Set Interface = Nothing
    Exit Sub
  End If

End Sub

