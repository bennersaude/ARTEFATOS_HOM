'HASH: 95F6E7F35A184EF8E53A5E324AE85843
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim sQL As Object

  Set sQL = NewQuery

  sQL.Add("SELECT MATMED FROM SAM_MATMEDPRECOTAB_VLR WHERE MATMEDPRECOTAB = " + CurrentQuery.FieldByName("MATMEDPRECOTAB").AsString)
  sQL.Active = True

  If Not sQL.EOF Then
    Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

    Linha = Interface.Vigencia(CurrentSystem, "SAM_MATMEDPRECOGERAL", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "MATMEDPRECOTAB", "")

    If Linha = "" Then
      CanContinue = True
    Else
      CanContinue = False
      bsShowMessage(Linha, "E")
      Exit Sub
    End If

  End If

End Sub

