'HASH: 9974E960ED34BC786C7FD22BA5BA4B10
Sub Main() 
  Dim Executar As Object 
  Dim Q        As Object
  Set Q = NewQuery
  Q.Add("SELECT USUARIOPADRAO, HOSTPADRAO FROM AEX_PARAMETROSGERAIS")
  Q.Active=True 
  Set Executar = CreateBennerObject("BSAte006.Rotinas")
  Executar.ProcessarAgendamento(CurrentSystem,"",346, _
                                Q.FieldByName("USUARIOPADRAO").AsInteger, Q.FieldByName("HOSTPADRAO").AsString)
  Set Executar = Nothing
  Set Q = Nothing
End Sub
