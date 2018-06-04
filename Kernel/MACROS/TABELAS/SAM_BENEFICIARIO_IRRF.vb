'HASH: BF3DD2C56C1988DF6EC09FE0220C97E6
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = " AND DATAFINAL < DATAINICIAL"

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_IRRF", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Set Interface = Nothing
    Exit Sub
  End If

End Sub

