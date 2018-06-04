'HASH: F214415C13445818D3DF742D4027D052
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)

'início sms 62791 - Edilson.Castro - 01/09/2006
  Dim q As Object
  Set q = NewQuery

  q.Add("SELECT 1")
  q.Add("  FROM SAM_MATMEDPRECOPRESTADOR")
  q.Add(" WHERE MATMEDPRECOTAB = :HandleTabela")
  q.ParamByName("HandleTabela").AsInteger = CurrentQuery.FieldByName("MATMEDPRECOTAB").AsInteger
  q.Active = True
  CanContinue = q.EOF
  q.Active = False

  Set q = Nothing
  If Not CanContinue Then
    bsShowMessage("Esta tabela já está vinculada a prestadores, não é possível cadastrar estados.", "E")
    Exit Sub
  End If
'fim sms 62791

  Dim Interface As Object
  Dim Linha As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_MATMEDPRECOESTADO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESTADO", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
End Sub
