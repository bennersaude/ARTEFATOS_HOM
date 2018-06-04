'HASH: FEDBBDF021F92E4D094AD974435D21B9
'Macro: SFN_FOLHAPAGTO_TIPOFATCODFOLHA
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT CFOLHA.HANDLE")
  SQL.Add("FROM SFN_FOLHAPAGTO_TIPOFAT TFAT")
  SQL.Add("JOIN SFN_FOLHAPAGTO_TIPOFATCODFOLHA CFOLHA ON CFOLHA.FOLHAPAGAMENTOTIPOFAT = TFAT.HANDLE")
  SQL.Add("WHERE TFAT.FOLHAPAGAMENTO = :HFOLHAPAGAMENTO")
  SQL.Add("  AND CFOLHA.CODIGOFOLHA = :HCODIGOFOLHA")
  SQL.Add("  AND CFOLHA.HANDLE <> :HATUAL")
  SQL.ParamByName("HFOLHAPAGAMENTO").Value = RecordHandleOfTable("SFN_FOLHAPAGTO")
  SQL.ParamByName("HCODIGOFOLHA").Value = CurrentQuery.FieldByName("CODIGOFOLHA").AsInteger
  SQL.ParamByName("HATUAL").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Este código folha já está cadastrado para esta folha de pagamento", "E")
    CanContinue = False
  End If

  Set SQL = Nothing
End Sub

