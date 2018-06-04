'HASH: B855B374619CD07231CBFA462ED17CC8
Sub Main() 
  Dim Executar As Object 
  Dim Q        As Object
  Set Q = NewQuery
  Q.Add("SELECT USUARIOPADRAO, HOSTPADRAO FROM AEX_PARAMETROSGERAIS")
  Q.Active=True 
  Set Executar = CreateBennerObject("BSAte006.Rotinas")
  Executar.ProcessarAgendamento(CurrentSystem,"",1, _
                                Q.FieldByName("USUARIOPADRAO").AsInteger, Q.FieldByName("HOSTPADRAO").AsString)
  Set Executar = Nothing
  Set Q = Nothing
End Sub
